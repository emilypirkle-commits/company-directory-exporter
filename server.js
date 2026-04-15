// ─────────────────────────────────────────────────────────────────────────────
// server.js  –  Company Directory Exporter
// ─────────────────────────────────────────────────────────────────────────────
// Two routes:
//   POST /scrape   – scrapes ALL pages of a directory, resolves company names,
//                    and returns the final dataset as JSON
//   POST /download – takes company data and returns a formatted Excel file
// ─────────────────────────────────────────────────────────────────────────────

const express = require('express');
const axios   = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');

const app  = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static('public'));


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 1 — SCRAPER CONFIGURATION
//  ─────────────────────────────────────────────────────────────────────────────
//  All the CSS selectors and tuning values live here.
//  If the scraper stops working on a site, this is the first place to edit.
// ═════════════════════════════════════════════════════════════════════════════

const SELECTORS = {

  // Each entry here is a CSS selector that should match ONE company card on the
  // directory page.  The scraper tries them in order and stops at the first hit.
  containers: [
    '.member-directory-listing',
    '.wpbdp-listing',
    '.geodir-category-listing',
    '.listing-item',
    '.directory-item',
    '.member-listing',
    '.member-item',
    '.member',
    '.company-item',
    '.grid-item',
    'article',
    'li.wpbdp-listing',
  ],

  // These selectors are tried INSIDE each container to find the company name.
  // They are used before the automatic fallback strategies.
  companyName: [
    '.member-name',
    '.listing-title',
    '.company-name',
    '.entry-title',
    '.wpbdp-field-display-value',
    'h1', 'h2', 'h3', 'h4', 'h5',
    'strong',
  ],

  // CSS selectors for the "next page" pagination link.
  nextPage: [
    'a.next',
    'a[rel="next"]',
    '.pagination a.next',
    '.nav-links a.next',
    'li.next a',
    '.next-page a',
    'a[aria-label="Next page"]',
    'a[aria-label="Next"]',
  ],
};

// Visible link text that means "go to the next page".
// Add more here if you encounter a site using a different label.
const NEXT_LINK_TEXT = ['next', '›', '»', 'next page', '>'];

// Stop after this many pages no matter what.
// Prevents runaway loops on broken or circular pagination.
const MAX_PAGES = 100;

// Wait this many milliseconds between page fetches.
// A small delay is polite and reduces the chance of being blocked.
const PAGE_DELAY_MS = 600;


// ── Company website resolution settings ───────────────────────────────────────

// How many company websites to fetch at the same time during name resolution.
// Increase for speed, decrease if you get connection errors or are blocked.
const RESOLVE_CONCURRENCY = 5;

// Give up on fetching a company's website after this many milliseconds.
// 5 seconds is generous for a homepage — slow sites beyond this are skipped.
// ▶ Raise this number if you find valid sites are being skipped too often.
const WEBSITE_FETCH_TIMEOUT_MS = 5000;

// A Set of lowercase strings that are too vague to be real company names.
// If a name matches one of these, we'll try the company website instead.
// ▶ To add more, put them inside the Set(...) below.
const GENERIC_NAMES = new Set([
  'home', 'homepage', 'welcome', 'untitled', 'index',
  'website', 'site', 'online', 'page', 'loading',
  'error', 'not found', '404',
]);

// Words/phrases that appear after separators in page <title> tags but are NOT
// part of the real company name.
// e.g.  "Acme Ltd | Home"        →  we keep "Acme Ltd", discard "Home"
// e.g.  "Acme - Official Website"→  we keep "Acme", discard "Official Website"
// ▶ Add more here if you see other junk patterns in the output.
const TITLE_JUNK_SEGMENTS = [
  'home', 'homepage', 'welcome', 'official site', 'official website',
  'the official site', 'the official website', 'website', 'online',
];


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 2 — GENERAL UTILITIES
// ═════════════════════════════════════════════════════════════════════════════

// Returns a promise that resolves after `ms` milliseconds.
// Used to add a polite pause between page requests.
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Fetches a URL and returns the raw HTML string.
// `timeoutMs` controls how long to wait before giving up.
async function fetchPage(url, timeoutMs = 20000) {
  const response = await axios.get(url, {
    headers: {
      // Pretend to be a regular browser so sites don't block us
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
// Examples that return true:  "https://acme.com", "www.acme.co.uk", "acme.co.uk"
// Examples that return false: "Acme Ltd", "2ManageProperty"
function looksLikeUrl(text) {
  return (
    /^https?:\/\//i.test(text) ||           // starts with http:// or https://
    /^www\./i.test(text) ||                  // starts with www.
    /^[a-z0-9-]+\.[a-z]{2,}(\/|$)/i.test(text) // bare domain like "acme.co.uk"
  );
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 3 — DIRECTORY PAGE EXTRACTION
//  (scraping company data from the listing pages)
// ═════════════════════════════════════════════════════════════════════════════

// ── extractName ────────────────────────────────────────────────────────────────
// Tries to pull a company name out of a single container element.
// Three strategies are tried in order:
//   1. Look for a specific CSS selector (e.g. .member-name, h3)
//   2. Find the first anchor link whose visible text is NOT a web address
//   3. Take the first line of text that doesn't look like a URL, phone, or email
function extractName($, el) {

  // Strategy 1: explicit CSS selectors from config
  for (const sel of SELECTORS.companyName) {
    const found = $(el).find(sel).first();
    const text  = found.text().trim();
    if (text) return text;
  }

  // Strategy 2: first anchor tag whose TEXT looks like a name (not a URL)
  // This catches patterns like:
  //   <a href="https://acme.com">Acme Ltd</a>          ← external link, text is name
  //   <a href="/members/acme-ltd/">Acme Ltd</a>         ← internal profile link
  let linkText = '';
  $(el).find('a[href]').each((_i, a) => {
    const href = ($(a).attr('href') || '').trim();
    const text = $(a).text().trim();

    // Skip links that are clearly not company names
    if (
      href.startsWith('tel:')        ||
      href.startsWith('mailto:')     ||
      href.startsWith('#')           ||
      href.startsWith('javascript:')
    ) return; // continue to next link

    // Skip if the visible text itself looks like a domain
    if (looksLikeUrl(text)) return;

    if (text.length > 1) {
      linkText = text;
      return false; // break — stop the loop
    }
  });
  if (linkText) return linkText;

  // Strategy 3: first meaningful line of text in the container.
  //
  // Important: Cheerio's .text() strips ALL tags silently, including <br>.
  // If the site uses <br> to separate fields like:
  //   CompanyName<br>Address<br>Phone
  // then .text() returns "CompanyNameAddressPhone" (all mashed together).
  //
  // Fix: grab the raw HTML, replace <br> tags with newlines BEFORE stripping tags.
  const rawHtml = $(el).html() || '';
  const textWithBreaks = rawHtml
    .replace(/<br\s*\/?>/gi, '\n') // <br> → real newline
    .replace(/<[^>]+>/g, ' ');     // all other tags → space

  const lines = textWithBreaks
    .split(/[\r\n]+/)
    .map(l => l.replace(/\s+/g, ' ').trim())
    .filter(l =>
      l.length > 1 &&
      !looksLikeUrl(l) &&
      !/^\+?[\d\s\-().]{6,}$/.test(l) && // not a phone number
      !l.includes('@')                     // not an email address
    );
  return lines[0] || '';
}

// ── extractWebsite ─────────────────────────────────────────────────────────────
// Returns the first external URL (http:// or https://) found in the container.
function extractWebsite($, el) {
  let website = '';
  $(el).find('a[href]').each((_i, a) => {
    const href = ($(a).attr('href') || '').trim();
    if (/^https?:\/\//i.test(href)) {
      website = href;
      return false; // break
    }
  });
  return website;
}

// ── findNextPageUrl ────────────────────────────────────────────────────────────
// Looks for a "next page" link on the current page.
// Returns the full URL of the next page, or null if there isn't one.
function findNextPageUrl($, currentUrl) {
  let nextHref = null;

  // Approach A: try CSS selectors (e.g. a.next, a[rel="next"])
  for (const sel of SELECTORS.nextPage) {
    const el = $(sel).first();
    if (el.length) {
      nextHref = el.attr('href') || null;
      if (nextHref) break;
    }
  }

  // Approach B: scan all links for text like "Next", "»", etc.
  if (!nextHref) {
    $('a[href]').each((_i, a) => {
      const text = $(a).text().trim().toLowerCase();
      if (NEXT_LINK_TEXT.includes(text)) {
        nextHref = $(a).attr('href') || null;
        if (nextHref) return false; // break
      }
    });
  }

  if (!nextHref) return null;

  // Resolve relative URLs against the current page
  // e.g. "?page=2" becomes "https://example.com/directory/?page=2"
  try {
    return new URL(nextHref, currentUrl).href;
  } catch {
    return null;
  }
}

// ── scrapeOnePage ──────────────────────────────────────────────────────────────
// Fetches one directory page URL and extracts all company entries on it.
// Returns: { companies: [{name, website}], nextUrl: string|null }
async function scrapeOnePage(url, pageNumber) {
  const html = await fetchPage(url);
  const $    = cheerio.load(html);

  // Find the first container selector that matches
  let containers      = $();
  let matchedSelector = null;

  for (const sel of SELECTORS.containers) {
    const found = $(sel);
    if (found.length > 0) {
      containers      = found;
      matchedSelector = sel;
      break;
    }
  }

  if (containers.length === 0) {
    console.log(`  ⚠  Page ${pageNumber}: no container selector matched`);
    return { companies: [], nextUrl: null };
  }

  console.log(`  ✓  Page ${pageNumber}: matched "${matchedSelector}" → ${containers.length} containers`);

  // Extract name + website from each container
  const companies = [];
  containers.each((_i, el) => {
    const name    = extractName($, el);
    const website = extractWebsite($, el);
    // Only keep entries where we found at least something
    if (name || website) {
      companies.push({ name, website });
    }
  });

  console.log(`  →  Page ${pageNumber}: extracted ${companies.length} raw entries`);

  const nextUrl = findNextPageUrl($, url);
  if (nextUrl) {
    console.log(`  ↪  Page ${pageNumber}: next page → ${nextUrl}`);
  } else {
    console.log(`  ✋  Page ${pageNumber}: no next page — stopping`);
  }

  return { companies, nextUrl };
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 4 — COMPANY NAME CLEANING & QUALITY CHECK
// ═════════════════════════════════════════════════════════════════════════════

// ── cleanName ──────────────────────────────────────────────────────────────────
// Strips common junk from a raw name string.
//
// Many page titles look like:  "Acme Ltd | Home"  or  "Acme - Official Website"
// This function splits on separators (|, -, –, —, :), throws away any segment
// that's in TITLE_JUNK_SEGMENTS, and returns the first remaining segment.
//
// ▶ If you see new junk patterns in the output, add them to TITLE_JUNK_SEGMENTS
//   near the top of this file.
//
// Examples:
//   "Acme Ltd | Home"           →  "Acme Ltd"
//   "Acme – Official Website"   →  "Acme"
//   "  Welcome  "               →  "Welcome"  (still caught by isWeakName)
function cleanName(raw) {
  if (!raw) return '';

  let name = raw.trim();

  // Split on pipe, dash variants, colon — these are separator characters
  const parts = name.split(/\s*[|\-–—:]\s*/);

  // Keep only parts that aren't in the junk list and aren't empty
  const meaningful = parts.filter(p => {
    const lower = p.trim().toLowerCase();
    return p.trim().length > 1 && !TITLE_JUNK_SEGMENTS.includes(lower);
  });

  // Take the first meaningful part, or fall back to the full original string
  name = (meaningful[0] || parts[0] || name).trim();

  // Remove any stray separator characters left at the start or end
  name = name.replace(/^[|\-–—:,.\s]+|[|\-–—:,.\s]+$/g, '').trim();

  return name;
}

// ── isWeakName ─────────────────────────────────────────────────────────────────
// Returns true if a name is NOT good enough to use — meaning we should try
// the company's own website to get a better name.
//
// ▶ If the quality check is too strict (rejecting valid names), loosen the rules.
// ▶ If it's too loose (passing bad names), add more conditions here.
function isWeakName(name, website) {
  if (!name || name.trim().length < 2) return true; // empty or single character

  const lower = name.trim().toLowerCase();

  // In our GENERIC_NAMES set (e.g. "home", "welcome", "website")
  if (GENERIC_NAMES.has(lower)) return true;

  // Too short — unlikely to be a real company name
  if (name.trim().length < 3) return true;

  // The text itself looks like a web address
  if (looksLikeUrl(name)) return true;

  // The name exactly matches the domain of the website
  // e.g. name = "acme", website = "https://acme.co.uk"
  if (website) {
    try {
      const hostname  = new URL(website).hostname.replace(/^www\./, '');
      const domainBase = hostname.split('.')[0].toLowerCase(); // "acme" from "acme.co.uk"
      if (lower === domainBase || lower === hostname.toLowerCase()) return true;
    } catch { /* ignore malformed URLs */ }
  }

  return false; // name looks fine
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 5 — COMPANY WEBSITE RESOLUTION
//  (visiting each company's own website to find a better name)
// ═════════════════════════════════════════════════════════════════════════════

// ── resolveNameFromWebsite ─────────────────────────────────────────────────────
// Visits a company's own website and tries to extract the company name.
//
// Sources tried in order (most reliable first):
//   1. <meta property="og:site_name">   — explicitly set by the website owner
//   2. <meta name="application-name">   — set by many CMS platforms
//   3. <title> tag                      — reliable but often has junk suffixes
//   4. Logo image alt text              — often the company name
//   5. <h1>                             — last resort; may be a tagline
//
// Returns: { name: string, source: string }  or  null if nothing usable found.
// Never throws — any network or parse error returns null silently.
async function resolveNameFromWebsite(url) {
  let html;
  try {
    html = await fetchPage(url, WEBSITE_FETCH_TIMEOUT_MS);
  } catch (err) {
    // Work out why it failed and log a clear one-line message.
    // axios puts the error type in err.code — ECONNABORTED means it timed out.
    const isTimeout = err.code === 'ECONNABORTED' || err.message?.includes('timeout');
    const reason    = isTimeout ? `timed out (>${WEBSITE_FETCH_TIMEOUT_MS}ms)` : err.message;
    console.log(`     ⏩ skipped ${url} — ${reason}`);
    return null; // carry on with the rest; this entry will be marked "Unknown"
  }

  const $ = cheerio.load(html);

  // ── 1. og:site_name ─────────────────────────────────────────────────────────
  // Example HTML: <meta property="og:site_name" content="Acme Ltd">
  const ogSiteName = ($('meta[property="og:site_name"]').attr('content') || '').trim();
  if (ogSiteName) {
    const cleaned = cleanName(ogSiteName);
    if (!isWeakName(cleaned, url)) return { name: cleaned, source: 'website_og_site_name' };
  }

  // ── 2. application-name ──────────────────────────────────────────────────────
  // Example HTML: <meta name="application-name" content="Acme Ltd">
  const appName = ($('meta[name="application-name"]').attr('content') || '').trim();
  if (appName) {
    const cleaned = cleanName(appName);
    if (!isWeakName(cleaned, url)) return { name: cleaned, source: 'website_application_name' };
  }

  // ── 3. <title> tag ───────────────────────────────────────────────────────────
  // Example HTML: <title>Acme Ltd | Home</title>
  // cleanName() strips the "| Home" part
  const title = ($('title').first().text() || '').trim();
  if (title) {
    const cleaned = cleanName(title);
    if (!isWeakName(cleaned, url)) return { name: cleaned, source: 'website_title' };
  }

  // ── 4. Logo image alt text ───────────────────────────────────────────────────
  // Looks for images likely to be the site logo based on class/id names
  const logoEl  = $('img[class*="logo"], img[id*="logo"], a[class*="logo"] img, header img').first();
  const logoAlt = (logoEl.attr('alt') || '').trim();
  if (logoAlt) {
    const cleaned = cleanName(logoAlt);
    if (!isWeakName(cleaned, url)) return { name: cleaned, source: 'website_logo_alt' };
  }

  // ── 5. <h1> tag ──────────────────────────────────────────────────────────────
  // Last resort — may be a tagline, page heading, or marketing copy
  const h1 = ($('h1').first().text() || '').trim();
  if (h1) {
    const cleaned = cleanName(h1);
    if (!isWeakName(cleaned, url)) return { name: cleaned, source: 'website_h1' };
  }

  return null; // nothing usable found on this website
}


// ── runWithConcurrency ─────────────────────────────────────────────────────────
// Runs an async function over an array of items, but limits how many can run
// at the same time.
//
// Think of it like a checkout queue: only `limit` people can be served at once.
// As soon as one finishes, the next one in the queue starts automatically.
//
// Why this matters: without a limit, we'd fire off hundreds of website requests
// simultaneously — likely triggering rate-limiting or crashing the server.
//
// Parameters:
//   items   — the list of things to process (any array)
//   asyncFn — an async function that processes one item: asyncFn(item, index)
//   limit   — maximum number of simultaneous operations
//
// Returns: an array of results in the same order as the input items.
async function runWithConcurrency(items, asyncFn, limit) {
  const results = new Array(items.length).fill(null);

  // Build a queue of tasks, each tagged with its original array index
  // so we can put results back in the right position
  const queue = items.map((item, i) => ({ item, i }));

  // Each worker pulls jobs from the shared queue until it's empty
  async function worker() {
    while (queue.length > 0) {
      const { item, i } = queue.shift(); // take next task
      results[i] = await asyncFn(item, i);
    }
  }

  // Start `limit` workers running in parallel — they all share the same queue
  const workers = Array.from(
    { length: Math.min(limit, items.length) },
    () => worker()
  );
  await Promise.all(workers);

  return results;
}


// ── resolveAllNames ────────────────────────────────────────────────────────────
// The main name resolution pipeline.
//
// Takes the raw { name, website } list from directory scraping and produces a
// clean { companyName, website, sourceOfName } list.
//
// Steps:
//   1. Check each directory-page name with isWeakName()
//   2. Collect all entries whose name is weak and that have a website URL
//   3. Fetch those websites concurrently (limited by RESOLVE_CONCURRENCY)
//   4. Cache results so the same domain is never fetched twice
//   5. Merge everything back into the final output array
async function resolveAllNames(rawCompanies) {

  // All entries go to the company website for name resolution.
  // The name scraped from the directory page is ignored — it was unreliable.
  // The website URL (scraped from the directory) is what we use to fetch the name.

  const withWebsite    = rawCompanies.filter(({ website }) => !!website);
  const withoutWebsite = rawCompanies.filter(({ website }) => !website);

  console.log(`\n  ── Name resolution ──────────────────────────────`);
  console.log(`     ${withWebsite.length} entries have a website — will fetch each one`);
  console.log(`     ${withoutWebsite.length} entries have no website — will be marked "Unknown"`);

  // Cache: normalised URL → resolved result ({ name, source }) or null.
  // Prevents fetching the same domain more than once if it appears in multiple entries.
  const websiteCache = new Map();

  // Build a deduplicated list of unique URLs to fetch
  const uniqueToFetch = [];
  for (const { website } of withWebsite) {
    const key = website.replace(/\/+$/, '').toLowerCase();
    if (!websiteCache.has(key)) {
      websiteCache.set(key, null); // placeholder so we don't add it twice
      uniqueToFetch.push({ key, url: website });
    }
  }

  console.log(`     (${uniqueToFetch.length} unique domains to fetch)`);

  // Fetch all unique websites concurrently, limited by RESOLVE_CONCURRENCY.
  // A counter lets us print "3/42" style progress in the terminal.
  let doneCount = 0;
  await runWithConcurrency(uniqueToFetch, async ({ key, url }) => {
    const result = await resolveNameFromWebsite(url);
    websiteCache.set(key, result);
    doneCount++;
    if (result) {
      console.log(`     ✓ [${doneCount}/${uniqueToFetch.length}] ${url} → "${result.name}" (${result.source})`);
    }
    // Timed-out/failed sites already logged their own ⏩ line inside resolveNameFromWebsite
  }, RESOLVE_CONCURRENCY);

  // Build the final output
  let resolvedCount = 0;
  let failedCount   = 0;

  const finalCompanies = rawCompanies.map(({ website }) => {
    if (website) {
      const key      = website.replace(/\/+$/, '').toLowerCase();
      const resolved = websiteCache.get(key);

      if (resolved && resolved.name) {
        resolvedCount++;
        return {
          companyName:  resolved.name,
          website:      website,
          sourceOfName: resolved.source,
        };
      }
    }

    // No website, or website fetch failed
    failedCount++;
    return {
      companyName:  'Unknown',
      website:      website,
      sourceOfName: 'unknown',
    };
  });

  console.log(`     ${resolvedCount} names resolved from company websites`);
  console.log(`     ${failedCount} could not be resolved → marked "Unknown"`);
  console.log(`  ─────────────────────────────────────────────────`);

  return finalCompanies;
}


// ═════════════════════════════════════════════════════════════════════════════
//  SECTION 6 — ROUTES
// ═════════════════════════════════════════════════════════════════════════════

// ── POST /scrape ───────────────────────────────────────────────────────────────
// Body:    { url: "https://..." }
// Returns: { companies: [{companyName, website, sourceOfName}], count, pageCount }
app.post('/scrape', async (req, res) => {
  const { url } = req.body;

  if (!url) {
    return res.status(400).json({ error: 'A URL is required.' });
  }

  let startUrl;
  try {
    startUrl = new URL(url).href; // normalise (e.g. encode spaces in query params)
  } catch {
    return res.status(400).json({ error: 'That does not look like a valid URL.' });
  }

  console.log(`\n${'─'.repeat(60)}`);
  console.log(`  Starting scrape: ${startUrl}`);
  console.log(`${'─'.repeat(60)}`);

  try {
    // ── Step 1: Scrape all directory pages ──────────────────────────────────
    const allRaw   = [];
    let currentUrl = startUrl;
    let pageNumber = 1;

    while (currentUrl && pageNumber <= MAX_PAGES) {
      const { companies, nextUrl } = await scrapeOnePage(currentUrl, pageNumber);
      allRaw.push(...companies);

      if (companies.length === 0) {
        console.log(`  ⚠  Page ${pageNumber} returned 0 entries — stopping pagination`);
        break;
      }

      currentUrl = nextUrl;
      pageNumber++;
      if (currentUrl) await sleep(PAGE_DELAY_MS);
    }

    if (pageNumber > MAX_PAGES) {
      console.log(`  ⚠  Reached MAX_PAGES limit (${MAX_PAGES}) — stopping`);
    }

    const totalPages = pageNumber - 1;
    console.log(`\n  Total pages scraped  : ${totalPages}`);
    console.log(`  Total raw entries    : ${allRaw.length}`);

    // ── Step 2: Deduplicate ─────────────────────────────────────────────────
    // Remove entries with the same website URL (or same name if no URL).
    // We deduplicate BEFORE website resolution so we don't waste fetch requests.
    const seenWebsites = new Set();
    const seenNames    = new Set();

    const deduped = allRaw.filter(({ name, website }) => {
      if (website) {
        const key = website.replace(/\/+$/, '').toLowerCase();
        if (seenWebsites.has(key)) return false;
        seenWebsites.add(key);
        return true;
      } else {
        const key = (name || '').toLowerCase();
        if (seenNames.has(key)) return false;
        seenNames.add(key);
        return true;
      }
    });

    console.log(`  After deduplication  : ${deduped.length} entries`);

    // ── Step 3: Resolve company names ───────────────────────────────────────
    // For entries with weak/missing names, visit the company website to find
    // a better name using og:site_name, title, etc.
    const companies = await resolveAllNames(deduped);

    console.log(`\n  Final company count  : ${companies.length}`);
    console.log(`${'─'.repeat(60)}\n`);

    if (companies.length === 0) {
      return res.status(422).json({
        error:
          'No companies were found. The page structure may not match any known ' +
          'selector, or the site may require JavaScript to render its content. ' +
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
// Body:    { companies: [{companyName, website, sourceOfName}] }
// Returns: a formatted .xlsx file
app.post('/download', async (req, res) => {
  const { companies } = req.body;

  if (!Array.isArray(companies) || companies.length === 0) {
    return res.status(400).json({ error: 'No company data provided.' });
  }

  try {
    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Companies');

    // Three columns: name, website, source
    worksheet.columns = [
      { header: 'Company Name', key: 'companyName',  width: 42 },
      { header: 'Website',      key: 'website',      width: 48 },
      { header: 'Name Source',  key: 'sourceOfName', width: 26 },
    ];

    // Header row: white bold text on blue background
    const headerRow = worksheet.getRow(1);
    headerRow.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    headerRow.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'left' };
    headerRow.height    = 22;

    // Add one row per company
    companies.forEach(({ companyName, website, sourceOfName }) => {
      const row = worksheet.addRow({ companyName, website, sourceOfName });

      // Make the website a clickable hyperlink in Excel
      if (website) {
        const cell = row.getCell('website');
        cell.value = { text: website, hyperlink: website };
        cell.font  = { color: { argb: 'FF2563EB' }, underline: true };
      }
    });

    // Zebra-stripe every other data row for readability
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && rowNumber % 2 === 0) {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
      }
    });

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
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
