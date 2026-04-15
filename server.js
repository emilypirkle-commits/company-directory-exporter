// ─────────────────────────────────────────────────────────────────────────────
// server.js  –  Company Directory Exporter
// ─────────────────────────────────────────────────────────────────────────────
// Routes:
//   GET  /        – serves the frontend (public/index.html)
//   POST /scrape  – scrapes ALL pages of a directory and returns JSON
//   POST /download – converts company data into a downloadable Excel file
// ─────────────────────────────────────────────────────────────────────────────

const express = require('express');
const axios   = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const path    = require('path');

const app  = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (_req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 1 — SITE PROFILES
//  ─────────────────────────────────────────────────────────────────────────────
//  A "site profile" tells the scraper exactly how a specific directory is built.
//  When the URL being scraped matches a known domain, that profile's settings
//  are used instead of (or before) the generic auto-detection logic.
//
//  HOW TO ADD A NEW SITE
//  ─────────────────────
//  1. Open the site in Chrome, right-click a company listing → Inspect
//  2. Find the element that wraps ONE complete listing and copy its class name
//  3. Add a new entry below with that class name in `containers`
//  4. Adjust `nameSelectors` and `websiteFilter` if needed
//
//  PROFILE FIELDS
//  ──────────────
//  layout          'card' | 'list' | 'searchResults'
//                  Describes the visual structure of the page.
//                  Used by auto-detection as a hint if no containers match.
//
//  containers      Array of CSS selectors tried in order.
//                  The first one that returns 2+ elements wins.
//                  ▶ Put the most specific/accurate selector first.
//
//  contentScope    (optional) CSS selector for the area of the page that
//                  contains the listings.  Used to avoid matching navigation
//                  or footer elements when containers are generic (e.g. 'li').
//
//  nameSelectors   (optional) CSS selectors tried inside each container to
//                  find the company name element.  Overrides the defaults.
//
//  websiteFilter   (optional) function(href) → boolean
//                  Return false to skip a link as the website URL.
//                  Use this to exclude internal links that look external.
//                  Example: href => !href.includes('yell.com')
//
//  nextPage        (optional) extra CSS selectors for the "next page" link.
//
//  nextPageText    (optional) extra visible-text strings meaning "next page".
// ═════════════════════════════════════════════════════════════════════════════

const SITE_PROFILES = {

  // ── The Property Institute ──────────────────────────────────────────────────
  'tpi.org.uk': {
    layout: 'card',
    containers: [
      '.member-directory-listing',
      '.member-listing',
      '.member-item',
      '.member',
      '.listing-item',
      'article',
    ],
  },

  // ── Thomson Local ───────────────────────────────────────────────────────────
  // Uses plain <li> elements with no CSS classes.
  // contentScope narrows the search to the main content area so we don't match
  // navigation or footer list items.
  'thomsonlocal.com': {
    layout: 'list',
    contentScope: 'main, #main, [id*="result"], [class*="result-list"], [class*="search-result"]',
    containers: ['li'],
    nameSelectors: ['h2', 'h3', 'h4'],
    // Exclude Thomson Local's own internal profile links
    websiteFilter: href => !href.includes('thomsonlocal.com'),
  },

  // ── Clutch.co ───────────────────────────────────────────────────────────────
  // ⚠ BLOCKED: Clutch uses Cloudflare. Standard HTTP requests return a
  // challenge page. This profile is ready for use with a headless browser
  // or residential proxy service.
  'clutch.co': {
    layout: 'card',
    containers: [
      '[class*="provider-list-item"]',
      '[class*="sg-provider"]',
      'li.provider',
      'article',
    ],
    nameSelectors: ['h3', 'h2', '[class*="company-name"]', '[class*="provider-name"]'],
  },

  // ── Yell.com ────────────────────────────────────────────────────────────────
  // ⚠ BLOCKED: Yell returns 403 for automated requests. This profile is ready
  // for use with a headless browser or residential proxy service.
  'yell.com': {
    layout: 'searchResults',
    containers: [
      '[class*="businessCapsule"]',
      '[class*="business-capsule"]',
      '[class*="businessCard"]',
      '[class*="listing"]',
      'article',
    ],
    nameSelectors: ['h2', 'h3', '[class*="businessName"]', '[class*="business-name"]'],
    // Exclude Yell's own internal profile links
    websiteFilter: href => !href.includes('yell.com'),
  },

  // ── Add more sites here following the same pattern ──────────────────────────
};


// ─────────────────────────────────────────────────────────────────────────────
//  FALLBACK SELECTORS
//  Used when no site profile matches, or when the profile's containers don't
//  find anything.  Organised by layout type so the most likely ones are tried
//  first.  You can add selectors here without touching individual profiles.
// ─────────────────────────────────────────────────────────────────────────────

const FALLBACK_SELECTORS = {
  // Card-based layouts — each listing is a styled box/card
  card: [
    '[class*="listing-item"]',
    '[class*="result-item"]',
    '[class*="business-card"]',
    '[class*="company-card"]',
    '[class*="member-card"]',
    '[class*="provider"]',
    '[class*="directory-item"]',
    '[class*="member"]',
    'article',
    '.card',
  ],

  // List-based layouts — each listing is a <li> inside a <ul> or <ol>
  list: [
    'main ul > li',
    '#content ul > li',
    '#results ul > li',
    '.results > ul > li',
    '[id*="result"] ul > li',
    '[class*="result"] ul > li',
    '[class*="listing"] > ul > li',
    'main ol > li',
  ],

  // Search-result layouts — varied structures typical of search pages
  searchResults: [
    '[class*="search-result"]',
    '[class*="serp-item"]',
    '[class*="result-card"]',
    '[class*="hit"]',   // Algolia-style search
    '[class*="listing"]',
    'article',
  ],

  // Last-resort selectors tried regardless of layout type
  generic: [
    'article',
    '[class*="item"]',
    '[class*="entry"]',
  ],
};

// Default name selectors tried inside any container (when no profile overrides them)
const DEFAULT_NAME_SELECTORS = [
  '.member-name', '.listing-title', '.company-name',
  '.business-name', '.entry-title', '.provider-name',
  'h1', 'h2', 'h3', 'h4', 'h5',
  'strong',
];

// Link text that signals "go to the next page"
const NEXT_PAGE_TEXT = ['next', '›', '»', 'next page', '>'];

// CSS selectors for "next page" links used when no profile specifies its own
const DEFAULT_NEXT_PAGE_SELECTORS = [
  'a.next', 'a[rel="next"]',
  '.pagination a.next', '.nav-links a.next',
  'li.next a', '.next-page a',
  'a[aria-label="Next page"]', 'a[aria-label="Next"]',
];

// Safety cap — stop after this many pages even if a "next" link keeps appearing
const MAX_PAGES = 100;

// Polite delay between page requests (milliseconds)
const PAGE_DELAY_MS = 600;

// Company website resolution settings
const RESOLVE_CONCURRENCY    = 5;     // max simultaneous website fetches
const WEBSITE_FETCH_TIMEOUT_MS = 5000; // ms before giving up on a slow site

// Names too vague to trust — triggers website resolution fallback
const GENERIC_NAMES = new Set([
  'home', 'homepage', 'welcome', 'untitled', 'index',
  'website', 'site', 'online', 'page', 'loading', 'error', 'not found', '404',
]);

// Junk suffixes stripped from page titles
const TITLE_JUNK_SEGMENTS = [
  'home', 'homepage', 'welcome', 'official site', 'official website',
  'the official site', 'the official website', 'website', 'online',
];


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 2 — PROFILE LOOKUP
// ═════════════════════════════════════════════════════════════════════════════

// Returns the SITE_PROFILE entry for a given URL, or null if none matches.
// Matches on the hostname so "www.yell.com" still matches the "yell.com" key.
function getProfile(url) {
  try {
    const hostname = new URL(url).hostname.replace(/^www\./, '');
    // Check for an exact domain match, then try parent domains
    // e.g. "london.thomsonlocal.com" would match "thomsonlocal.com"
    for (const domain of Object.keys(SITE_PROFILES)) {
      if (hostname === domain || hostname.endsWith('.' + domain)) {
        return SITE_PROFILES[domain];
      }
    }
  } catch { /* ignore malformed URLs */ }
  return null;
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 3 — GENERAL UTILITIES
// ═════════════════════════════════════════════════════════════════════════════

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Fetches a URL and returns the raw HTML string.
async function fetchPage(url, timeoutMs = 20000) {
  const response = await axios.get(url, {
    headers: {
      'User-Agent':
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
        'AppleWebKit/537.36 (KHTML, like Gecko) ' +
        'Chrome/124.0.0.0 Safari/537.36',
      'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'en-GB,en;q=0.5',
    },
    timeout: timeoutMs,
  });
  return response.data;
}

// Returns true if `text` looks like a web address rather than a company name.
function looksLikeUrl(text) {
  return (
    /^https?:\/\//i.test(text) ||
    /^www\./i.test(text) ||
    /^[a-z0-9-]+\.[a-z]{2,}(\/|$)/i.test(text)
  );
}

// Returns true if an element is inside a nav, header, footer, or sidebar.
// Used to discard list items that are part of the page chrome rather than listings.
function isPageChrome($, el) {
  return $(el).closest('nav, header, footer, aside, [class*="nav"], [class*="menu"], [class*="sidebar"]').length > 0;
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 4 — CONTAINER DETECTION
//  ─────────────────────────────────────────────────────────────────────────────
//  Finding the right "container" — the element that wraps one complete listing
//  — is the hardest part of scraping any directory.
//
//  This section tries selectors in a specific priority order:
//    1. Site profile containers (most accurate, site-specific)
//    2. Layout-appropriate fallbacks from FALLBACK_SELECTORS
//    3. Generic last-resort selectors
//
//  Each candidate is scored: it must match at least 2 elements and must not
//  consist entirely of page-chrome elements (nav, header, footer).
// ═════════════════════════════════════════════════════════════════════════════

// ── getSearchRoot ─────────────────────────────────────────────────────────────
// If the profile has a contentScope, narrow the search to that element.
// This is critical for list layouts where <li> appears in nav/footer too.
function getSearchRoot($, profile) {
  if (profile && profile.contentScope) {
    const scopeEl = $(profile.contentScope).first();
    if (scopeEl.length) {
      console.log(`     🔍 Content scoped to: "${profile.contentScope}"`);
      return scopeEl;
    }
  }
  return null; // null means "search the whole page"
}

// ── trySelector ───────────────────────────────────────────────────────────────
// Runs a CSS selector against either a scoped root element or the full page.
// Returns a filtered Cheerio set that excludes page-chrome elements,
// or an empty set if the selector returned fewer than 2 real results.
function trySelector($, selector, scopeRoot) {
  // If we have a scoped root element, search inside it; otherwise search the page
  const raw = scopeRoot ? scopeRoot.find(selector) : $(selector);

  // Filter out anything inside nav / header / footer
  const filtered = raw.filter((_i, el) => !isPageChrome($, el));

  // Require at least 2 matches — a single match is probably not a listing grid
  return filtered.length >= 2 ? filtered : $();
}

// ── findContainers ─────────────────────────────────────────────────────────────
// Main entry point for container detection.
// Returns { containers: CheerioSet, matchedSelector: string, method: string }
function findContainers($, profile) {
  const scopeRoot = getSearchRoot($, profile);
  const layout    = (profile && profile.layout) || 'card'; // default to card

  // ── Priority 1: profile-specific containers ──────────────────────────────
  if (profile && profile.containers) {
    for (const sel of profile.containers) {
      const found = trySelector($, sel, scopeRoot);
      if (found.length) {
        return { containers: found, matchedSelector: sel, method: 'profile' };
      }
    }
  }

  // ── Priority 2: layout-appropriate fallbacks ─────────────────────────────
  // Use the layout type as a hint to pick the most likely selectors first
  const layoutSelectors  = FALLBACK_SELECTORS[layout]  || [];
  const genericSelectors = FALLBACK_SELECTORS.generic   || [];
  const allFallbacks     = [...layoutSelectors, ...genericSelectors];

  // Also add the other layout types as a last attempt, so we don't give up
  // if the layout detection was wrong
  for (const otherLayout of ['card', 'list', 'searchResults']) {
    if (otherLayout !== layout) {
      allFallbacks.push(...(FALLBACK_SELECTORS[otherLayout] || []));
    }
  }

  for (const sel of allFallbacks) {
    const found = trySelector($, sel, scopeRoot);
    if (found.length) {
      return { containers: found, matchedSelector: sel, method: 'fallback' };
    }
  }

  // ── Nothing found ────────────────────────────────────────────────────────
  return { containers: $(), matchedSelector: null, method: 'none' };
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 5 — DATA EXTRACTION (NAME + WEBSITE FROM A SINGLE CONTAINER)
// ═════════════════════════════════════════════════════════════════════════════

// ── extractName ───────────────────────────────────────────────────────────────
// Tries three strategies in order to get the company name from one container.
// Accepts `nameSelectors` from the site profile (or falls back to defaults).
function extractName($, el, nameSelectors) {
  const selectors = nameSelectors || DEFAULT_NAME_SELECTORS;

  // Strategy 1: explicit CSS selectors (profile or defaults)
  for (const sel of selectors) {
    const text = $(el).find(sel).first().text().trim();
    if (text) return text;
  }

  // Strategy 2: first anchor whose visible text is NOT a web address
  // Catches both internal profile links and external links where the
  // link text is the company name (e.g. <a href="https://acme.com">Acme Ltd</a>)
  let linkText = '';
  $(el).find('a[href]').each((_i, a) => {
    const href = ($(a).attr('href') || '').trim();
    const text = $(a).text().trim();
    if (href.startsWith('tel:') || href.startsWith('mailto:') ||
        href.startsWith('#')    || href.startsWith('javascript:')) return;
    if (looksLikeUrl(text)) return;
    if (text.length > 1) { linkText = text; return false; } // break
  });
  if (linkText) return linkText;

  // Strategy 3: first meaningful line of container text.
  // Replace <br> with newlines before stripping tags so words don't run together.
  const rawHtml = $(el).html() || '';
  const lines = rawHtml
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .split(/[\r\n]+/)
    .map(l => l.replace(/\s+/g, ' ').trim())
    .filter(l => l.length > 1 && !looksLikeUrl(l) &&
                 !/^\+?[\d\s\-().]{6,}$/.test(l) && !l.includes('@'));
  return lines[0] || '';
}

// ── extractWebsite ─────────────────────────────────────────────────────────────
// Returns the first external URL in the container that passes the profile's
// websiteFilter (if one is defined).
function extractWebsite($, el, profile) {
  const filter = profile && profile.websiteFilter;
  let website = '';
  $(el).find('a[href]').each((_i, a) => {
    const href = ($(a).attr('href') || '').trim();
    if (!/^https?:\/\//i.test(href)) return;      // must be a full URL
    if (filter && !filter(href)) return;           // must pass the profile filter
    website = href;
    return false; // break
  });
  return website;
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 6 — PAGINATION
// ═════════════════════════════════════════════════════════════════════════════

// Returns the URL of the next page, or null if there isn't one.
// Combines profile-specific selectors with the universal defaults.
function findNextPageUrl($, currentUrl, profile) {
  const extraSelectors = (profile && profile.nextPage)   || [];
  const extraText      = (profile && profile.nextPageText)|| [];
  const allSelectors   = [...extraSelectors, ...DEFAULT_NEXT_PAGE_SELECTORS];
  const allText        = [...extraText, ...NEXT_PAGE_TEXT];

  let nextHref = null;

  // Approach A: CSS selectors
  for (const sel of allSelectors) {
    const href = $(sel).first().attr('href');
    if (href) { nextHref = href; break; }
  }

  // Approach B: scan visible link text
  if (!nextHref) {
    $('a[href]').each((_i, a) => {
      if (allText.includes($(a).text().trim().toLowerCase())) {
        nextHref = $(a).attr('href');
        if (nextHref) return false; // break
      }
    });
  }

  if (!nextHref) return null;

  try { return new URL(nextHref, currentUrl).href; }
  catch { return null; }
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 7 — PAGE SCRAPER
// ═════════════════════════════════════════════════════════════════════════════

// Fetches one directory page and returns all company entries found on it.
async function scrapeOnePage(url, pageNumber, profile) {
  const html = await fetchPage(url);
  const $    = cheerio.load(html);

  const { containers, matchedSelector, method } = findContainers($, profile);

  if (!containers.length) {
    console.log(`  ⚠  Page ${pageNumber}: no containers found (tried profile + all fallbacks)`);
    return { companies: [], nextUrl: null };
  }

  const methodLabel = method === 'profile' ? '📋 profile' : '🔍 auto-detected';
  console.log(`  ✓  Page ${pageNumber}: ${methodLabel} selector "${matchedSelector}" → ${containers.length} containers`);

  // Extract name + website from each container
  const nameSelectors = profile && profile.nameSelectors;
  const companies = [];

  containers.each((_i, el) => {
    const name    = extractName($, el, nameSelectors);
    const website = extractWebsite($, el, profile);
    if (name || website) companies.push({ name, website });
  });

  console.log(`  →  Page ${pageNumber}: extracted ${companies.length} raw entries`);

  const nextUrl = findNextPageUrl($, url, profile);
  console.log(nextUrl
    ? `  ↪  Page ${pageNumber}: next page → ${nextUrl}`
    : `  ✋  Page ${pageNumber}: no next page — stopping`);

  return { companies, nextUrl };
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 8 — COMPANY NAME CLEANING & QUALITY CHECK
// ═════════════════════════════════════════════════════════════════════════════

// Strips junk suffixes like "| Home" or "- Official Website" from a raw string.
// ▶ Add more patterns to TITLE_JUNK_SEGMENTS at the top if needed.
function cleanName(raw) {
  if (!raw) return '';
  let name = raw.trim();
  const parts = name.split(/\s*[|\-–—:]\s*/);
  const meaningful = parts.filter(p =>
    p.trim().length > 1 && !TITLE_JUNK_SEGMENTS.includes(p.trim().toLowerCase())
  );
  name = (meaningful[0] || parts[0] || name).trim();
  return name.replace(/^[|\-–—:,.\s]+|[|\-–—:,.\s]+$/g, '').trim();
}

// Returns true if a name is too weak to use — triggers website resolution.
// ▶ Adjust rules here if valid names are being rejected or junk is getting through.
function isWeakName(name, website) {
  if (!name || name.trim().length < 2) return true;
  const lower = name.trim().toLowerCase();
  if (GENERIC_NAMES.has(lower)) return true;
  if (name.trim().length < 3)   return true;
  if (looksLikeUrl(name))       return true;
  if (website) {
    try {
      const domain = new URL(website).hostname.replace(/^www\./, '').split('.')[0];
      if (lower === domain.toLowerCase()) return true;
    } catch { /* ignore */ }
  }
  return false;
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 9 — COMPANY WEBSITE RESOLUTION
//  (visiting each company's own website to find a reliable name)
// ═════════════════════════════════════════════════════════════════════════════

// Visits a company website and tries to extract the company name.
// Sources tried in order: og:site_name → application-name → <title>
//                         → logo alt text → <h1>
// Returns { name, source } or null on failure.
async function resolveNameFromWebsite(url) {
  let html;
  try {
    html = await fetchPage(url, WEBSITE_FETCH_TIMEOUT_MS);
  } catch (err) {
    const isTimeout = err.code === 'ECONNABORTED' || err.message?.includes('timeout');
    console.log(`     ⏩ skipped ${url} — ${isTimeout ? `timed out (>${WEBSITE_FETCH_TIMEOUT_MS}ms)` : err.message}`);
    return null;
  }

  const $ = cheerio.load(html);
  const sources = [
    { sel: () => $('meta[property="og:site_name"]').attr('content'), key: 'website_og_site_name' },
    { sel: () => $('meta[name="application-name"]').attr('content'), key: 'website_application_name' },
    { sel: () => $('title').first().text(),                          key: 'website_title' },
    { sel: () => $('img[class*="logo"],img[id*="logo"],a[class*="logo"] img,header img').first().attr('alt'), key: 'website_logo_alt' },
    { sel: () => $('h1').first().text(),                             key: 'website_h1' },
  ];

  for (const { sel, key } of sources) {
    const raw     = (sel() || '').trim();
    const cleaned = cleanName(raw);
    if (!isWeakName(cleaned, url)) return { name: cleaned, source: key };
  }
  return null;
}

// Limits simultaneous async operations to `limit` at a time.
// Works like a checkout queue — as one finishes, the next starts.
async function runWithConcurrency(items, asyncFn, limit) {
  const results = new Array(items.length).fill(null);
  const queue   = items.map((item, i) => ({ item, i }));
  async function worker() {
    while (queue.length) {
      const { item, i } = queue.shift();
      results[i] = await asyncFn(item, i);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

// Visits every company website to resolve a reliable name.
// Directory-page names are ignored entirely — too unreliable across sites.
async function resolveAllNames(rawCompanies) {
  const withWebsite    = rawCompanies.filter(c => !!c.website);
  const withoutWebsite = rawCompanies.filter(c => !c.website);

  console.log(`\n  ── Name resolution ───────────────────────────────`);
  console.log(`     ${withWebsite.length} entries have a website URL`);
  console.log(`     ${withoutWebsite.length} have no website — will be "Unknown"`);

  // Build a deduplicated list of URLs to fetch
  const websiteCache  = new Map();
  const uniqueToFetch = [];
  for (const { website } of withWebsite) {
    const key = website.replace(/\/+$/, '').toLowerCase();
    if (!websiteCache.has(key)) {
      websiteCache.set(key, null);
      uniqueToFetch.push({ key, url: website });
    }
  }
  console.log(`     (${uniqueToFetch.length} unique domains to fetch)`);

  let doneCount = 0;
  await runWithConcurrency(uniqueToFetch, async ({ key, url }) => {
    const result = await resolveNameFromWebsite(url);
    websiteCache.set(key, result);
    doneCount++;
    if (result) {
      console.log(`     ✓ [${doneCount}/${uniqueToFetch.length}] ${url} → "${result.name}" (${result.source})`);
    }
  }, RESOLVE_CONCURRENCY);

  let resolvedCount = 0;
  let failedCount   = 0;

  const finalCompanies = rawCompanies.map(({ website }) => {
    if (website) {
      const key      = website.replace(/\/+$/, '').toLowerCase();
      const resolved = websiteCache.get(key);
      if (resolved?.name) {
        resolvedCount++;
        return { companyName: resolved.name, website, sourceOfName: resolved.source };
      }
    }
    failedCount++;
    return { companyName: 'Unknown', website, sourceOfName: 'unknown' };
  });

  console.log(`     ${resolvedCount} resolved · ${failedCount} unknown`);
  console.log(`  ─────────────────────────────────────────────────`);
  return finalCompanies;
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 10 — ROUTES
// ═════════════════════════════════════════════════════════════════════════════

// ── POST /scrape ───────────────────────────────────────────────────────────────
app.post('/scrape', async (req, res) => {
  const { url } = req.body;
  if (!url) return res.status(400).json({ error: 'A URL is required.' });

  let startUrl;
  try { startUrl = new URL(url).href; }
  catch { return res.status(400).json({ error: 'That does not look like a valid URL.' }); }

  // Look up a site profile for this domain
  const profile = getProfile(startUrl);

  console.log(`\n${'─'.repeat(60)}`);
  console.log(`  Scraping: ${startUrl}`);
  console.log(`  Profile:  ${profile ? 'matched (' + new URL(startUrl).hostname + ')' : 'none — using auto-detection'}`);
  if (profile) console.log(`  Layout:   ${profile.layout}`);
  console.log(`${'─'.repeat(60)}`);

  try {
    // ── Scrape all pages ──────────────────────────────────────────────────
    const allRaw   = [];
    let currentUrl = startUrl;
    let pageNumber = 1;

    while (currentUrl && pageNumber <= MAX_PAGES) {
      const { companies, nextUrl } = await scrapeOnePage(currentUrl, pageNumber, profile);
      allRaw.push(...companies);
      if (!companies.length) {
        console.log(`  ⚠  Page ${pageNumber} returned 0 entries — stopping`);
        break;
      }
      currentUrl = nextUrl;
      pageNumber++;
      if (currentUrl) await sleep(PAGE_DELAY_MS);
    }

    const totalPages = pageNumber - 1;
    console.log(`\n  Pages scraped  : ${totalPages}`);
    console.log(`  Raw entries    : ${allRaw.length}`);

    // ── Deduplicate ───────────────────────────────────────────────────────
    const seenWeb   = new Set();
    const seenNames = new Set();
    const deduped   = allRaw.filter(({ name, website }) => {
      if (website) {
        const key = website.replace(/\/+$/, '').toLowerCase();
        if (seenWeb.has(key)) return false;
        seenWeb.add(key); return true;
      } else {
        const key = (name || '').toLowerCase();
        if (seenNames.has(key)) return false;
        seenNames.add(key); return true;
      }
    });
    console.log(`  After dedup    : ${deduped.length}`);

    // ── Resolve names from company websites ───────────────────────────────
    const companies = await resolveAllNames(deduped);
    console.log(`\n  Final count    : ${companies.length}`);
    console.log(`${'─'.repeat(60)}\n`);

    if (!companies.length) {
      return res.status(422).json({
        error:
          'No companies found. The page may use JavaScript to render its content ' +
          '(try a headless browser), or no selector matched the layout. ' +
          'Check the server terminal for details.',
      });
    }

    res.json({ companies, count: companies.length, pageCount: totalPages });

  } catch (err) {
    console.error('Scrape error:', err.message);
    res.status(500).json({ error: `Could not fetch the page: ${err.message}` });
  }
});


// ── POST /download ─────────────────────────────────────────────────────────────
app.post('/download', async (req, res) => {
  const { companies } = req.body;
  if (!Array.isArray(companies) || !companies.length) {
    return res.status(400).json({ error: 'No company data provided.' });
  }

  try {
    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Companies');

    worksheet.columns = [
      { header: 'Company Name', key: 'companyName',  width: 42 },
      { header: 'Website',      key: 'website',      width: 48 },
      { header: 'Name Source',  key: 'sourceOfName', width: 26 },
    ];

    const hRow = worksheet.getRow(1);
    hRow.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    hRow.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    hRow.alignment = { vertical: 'middle', horizontal: 'left' };
    hRow.height    = 22;

    companies.forEach(({ companyName, website, sourceOfName }) => {
      const row = worksheet.addRow({ companyName, website, sourceOfName });
      if (website) {
        const cell = row.getCell('website');
        cell.value = { text: website, hyperlink: website };
        cell.font  = { color: { argb: 'FF2563EB' }, underline: true };
      }
    });

    worksheet.eachRow((row, n) => {
      if (n > 1 && n % 2 === 0)
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="companies.xlsx"');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('Download error:', err.message);
    res.status(500).json({ error: `Could not create Excel file: ${err.message}` });
  }
});


// ── Start ──────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n  Company Directory Exporter`);
  console.log(`  Running at http://localhost:${PORT}\n`);
});
