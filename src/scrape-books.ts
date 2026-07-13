/**
 * scrape-books.ts — Collect book/chapter DOIs from the live topic listings.
 *
 * In short: the "book" listings on pharmacylibrary.com (PharmacotherapyFirst, the
 * OTC topics, the technician series, the media library, …) are paginated lists of
 * entries, each with its own DOI. This script walks every page of each configured
 * list, reads the DOI + title of each entry, and folds them into the "Books" tab
 * of Migration_record.xlsx — ONE row per DOI, tagged with the source's Category.
 *
 * Reads : each source URL (paginated) + the existing Books tab (for dedupe).
 * Writes: to the "Books" sheet of Migration_record.xlsx. A new DOI is appended;
 *         a DOI already present just gains the new Category in its comma-separated
 *         Category cell (never a second row). Existing Category cells are re-cased
 *         to canonical PascalCase on load. Other sheet(s) are left untouched.
 *
 * The listing links + their categories live in the SOURCES map in this file, keyed
 * by a slug. BOOK_TYPE selects which to scrape (a slug, a comma list, or "all").
 */

import * as dotenv from 'dotenv';
import * as fs     from 'fs';
import * as path   from 'path';
import * as XLSX   from 'xlsx';
import { chromium, BrowserContext, Browser, Page } from 'playwright';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const ROOT       = path.resolve(__dirname, '..');
const XLSX_FILE  = path.join(ROOT, 'Data', 'DOIs record', 'Migration_record.xlsx');
const SHEET_NAME = 'Books';

// Column order of the Books sheet (must match the existing header row).
const COLUMNS = ['DOI', 'Name', 'Status from Atypon', 'Date of attempt', 'Category'];

const BASE_URL   = 'https://pharmacylibrary.com';
const PAGE_TIMEOUT = 60_000;
const PAGE_SIZE  = parseInt(process.env.BOOK_PAGE_SIZE ?? '20', 10);
const MAX_PAGES  = parseInt(process.env.BOOK_MAX_PAGES ?? '1000', 10);   // safety cap
const DELAY_BETWEEN_PAGES = parseInt(process.env.BOOK_DELAY_MS ?? '800', 10);

// Each entry in the list of chapters/topics on a source page.
const ITEM_SELECTOR    = 'li.search__item';
const CONTENT_TIMEOUT  = parseInt(process.env.BOOK_CONTENT_TIMEOUT ?? '20000', 10);

// A "source" is one paginated topic listing plus the Category to tag its DOIs.
interface Source { url: string; category: string; }

// The topic listing links live HERE in the script (not in .env). Each is keyed
// by a `type` slug that BOOK_TYPE selects. Add a new listing by adding an entry.
const SOURCES: Record<string, Source> = {
  books:             { url: `${BASE_URL}/action/showPublications`,        category: 'Books'             },
  culturaltoolkit:   { url: `${BASE_URL}/topic/culturaltoolkit`,          category: 'CulturalToolkit'   },
  technicianseries:  { url: `${BASE_URL}/series/aphapts`,                 category: 'TechnicianSeries'  },
  otctopic:          { url: `${BASE_URL}/topic/aphaotctopics`,            category: 'OtcTopic'          },
  otccases:          { url: `${BASE_URL}/topic/aphaotccases`,             category: 'OtcCases'          },
  otcdecisiontrees:  { url: `${BASE_URL}/topic/aphaotcdecisiontrees`,     category: 'OtcDecisionTrees'  },
  otcfaculty:        { url: `${BASE_URL}/topic/aphaotcfaculty`,           category: 'OtcFaculty'        },
  pfirstmodules:     { url: `${BASE_URL}/topic/pfdsc`,                    category: 'PfirstModules'     },
  pfirstcases:       { url: `${BASE_URL}/topic/pfcbl`,                    category: 'PfirstCases'       },
  pfirstfaculty:     { url: `${BASE_URL}/topic/p1faculty`,                category: 'PfirstFaculty'     },
  medialibrary:      { url: `${BASE_URL}/topic/multimedia`,               category: 'MediaLibrary'      },
};

// BOOK_TYPE selects which listing(s) to scrape by their key in SOURCES above
// (e.g. "pfirst", "otc", "technicianseries", "otccases", ...), a comma list like
// "pfirst,otc", or "all"/blank for every source. Unknown types error out.
function resolveSources(): Source[] {
  const raw = (process.env.BOOK_TYPE ?? '').trim().toLowerCase();
  if (!raw || raw === 'all') return Object.values(SOURCES);
  const keys = raw.split(',').map(s => s.trim()).filter(Boolean);
  const unknown = keys.filter(k => !(k in SOURCES));
  if (unknown.length) {
    throw new Error(
      `Unknown BOOK_TYPE value(s): ${unknown.join(', ')}. ` +
      `Valid types: ${Object.keys(SOURCES).join(', ')}, all.`
    );
  }
  return keys.map(k => SOURCES[k]);
}

// ---------------------------------------------------------------------------
// Auth (cookie/login) — same approach as scrape-questions.ts / find-dois.ts
// ---------------------------------------------------------------------------

async function handleCloudflareChallenge(page: Page): Promise<void> {
  await page.waitForTimeout(2000);
  const cfFrame = page.frames().find(f => f.url().includes('challenges.cloudflare.com'));
  if (!cfFrame) {
    const directChallenge = await page.$('#content .cb-c, #content input[type="checkbox"]');
    if (!directChallenge) return;
    console.log('[Auth] Cloudflare Turnstile detected (direct) — clicking checkbox...');
    await page.click('#content input[type="checkbox"]').catch(() => {});
  } else {
    console.log('[Auth] Cloudflare Turnstile detected (iframe) — clicking checkbox...');
    try {
      const checkbox = cfFrame.locator('input[type="checkbox"]');
      await checkbox.waitFor({ state: 'attached', timeout: 8000 });
      await checkbox.click();
    } catch {
      console.warn('[Auth] Could not click Turnstile checkbox — may require manual intervention');
      return;
    }
  }
  console.log('[Auth] Waiting for Turnstile verification...');
  try {
    await Promise.race([
      page.waitForURL(url => !url.href.includes('challenges.cloudflare.com'), { timeout: 20000 }),
      page.waitForFunction(
        () => {
          const s = document.querySelector('#success') as HTMLElement | null;
          return s !== null && s.style.display !== 'none' && s.style.visibility !== 'hidden';
        },
        { timeout: 20000 }
      ),
    ]);
    console.log('[Auth] Turnstile challenge passed.');
    await page.waitForTimeout(1500);
  } catch {
    console.warn('[Auth] Turnstile did not resolve within 20 s.');
  }
}

async function applyAuth(context: BrowserContext): Promise<void> {
  const authMethod = (process.env.AUTH_METHOD ?? 'login').trim().toLowerCase();

  if (authMethod === 'cookie') {
    type CookieParam = {
      name: string; value: string; domain: string; path: string;
      httpOnly: boolean; secure: boolean; sameSite: 'Strict' | 'Lax' | 'None';
    };
    const cookies: CookieParam[] = [];
    if (process.env.CF_CLEARANCE) {
      cookies.push({
        name: 'cf_clearance', value: process.env.CF_CLEARANCE,
        domain: '.pharmacylibrary.com', path: '/',
        httpOnly: false, secure: true, sameSite: 'None',
      });
    }
    if (process.env.SESSION_COOKIE_VALUE) {
      cookies.push({
        name: process.env.SESSION_COOKIE_NAME ?? 'JSESSIONID',
        value: process.env.SESSION_COOKIE_VALUE,
        domain: '.pharmacylibrary.com', path: '/',
        httpOnly: true, secure: true, sameSite: 'Lax',
      });
    }
    if (cookies.length === 0) {
      console.warn('[Auth] AUTH_METHOD=cookie but no CF_CLEARANCE or SESSION_COOKIE_VALUE set in .env');
    } else {
      await context.addCookies(cookies);
      console.log(`[Auth] Injected ${cookies.length} cookie(s) — skipping login form`);
    }
    return;
  }

  const loginUrl = process.env.AUTH_LOGIN_URL ?? 'https://pharmacylibrary.com/action/login';
  const username = process.env.AUTH_USERNAME ?? '';
  const password = process.env.AUTH_PASSWORD ?? '';
  if (!username || !password) {
    console.warn('[Auth] WARNING: AUTH_USERNAME or AUTH_PASSWORD is empty — pages may return 403');
    return;
  }
  console.log(`[Auth] Navigating to login page: ${loginUrl}`);
  const loginPage = await context.newPage();
  try {
    await loginPage.goto(loginUrl, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT });
    await handleCloudflareChallenge(loginPage);
    await loginPage.waitForSelector(
      'input[name="login"], input[name="email"], input[type="email"]', { timeout: 10000 }
    ).catch(() => {});
    const usernameField = await loginPage.$('input[name="login"]') ??
                          await loginPage.$('input[name="email"]') ??
                          await loginPage.$('input[type="email"]');
    if (usernameField) await usernameField.fill(username);
    const passwordField = await loginPage.$('input[name="password"]') ??
                          await loginPage.$('input[type="password"]');
    if (passwordField) await passwordField.fill(password);
    await loginPage.click('button[type="submit"], input[type="submit"]');
    await loginPage.waitForNavigation({ waitUntil: 'networkidle', timeout: PAGE_TIMEOUT }).catch(() => {});
    console.log(`[Auth] Login flow finished → ${loginPage.url()}`);
  } finally {
    await loginPage.close();
  }
}

// ---------------------------------------------------------------------------
// Cloudflare challenge helpers
// ---------------------------------------------------------------------------

async function isCloudflareChallenge(page: Page): Promise<boolean> {
  try {
    const title = (await page.title()).toLowerCase();
    if (
      title.includes('security verification') ||
      title.includes('just a moment') ||
      title.includes('attention required')
    ) return true;
    const bodySnippet = await page.evaluate(() => (document.body?.innerText ?? '').slice(0, 500));
    return (
      bodySnippet.includes('Performing security verification') ||
      bodySnippet.includes('security service to protect against malicious bots') ||
      bodySnippet.includes('Verify you are human')
    );
  } catch {
    return false;
  }
}

async function waitForChallengeAndReload(page: Page, url: string): Promise<number> {
  console.warn('\n[CF] Challenge detected — waiting up to 30 s for auto-resolve...');
  await page.waitForFunction(
    () => {
      const t = document.title.toLowerCase();
      return (
        !t.includes('security verification') &&
        !t.includes('just a moment') &&
        !t.includes('attention required')
      );
    },
    { timeout: 30_000 }
  ).catch(() => {});
  await page.waitForTimeout(1500);
  const response = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT }).catch(() => null);
  return response?.status() ?? 0;
}

async function gotoWithRetry(page: Page, url: string, maxAttempts = 3, baseDelayMs = 2_000): Promise<number> {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp   = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT });
      const status = resp?.status() ?? 0;
      if (status === 404) return 404;
      if (status >= 200 && status < 300) return status;
      if (attempt < maxAttempts) await page.waitForTimeout(baseDelayMs * attempt);
      else return status;
    } catch (err: unknown) {
      const msg = (err as Error).message ?? '';
      if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) throw err;
      if (attempt < maxAttempts) await page.waitForTimeout(baseDelayMs * attempt);
      else return 0;
    }
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Book entry + page scraping
// ---------------------------------------------------------------------------

// Build a page URL for a given zero-based page index.
function pageUrl(baseUrl: string, pageIdx: number): string {
  const u = new URL(baseUrl);
  u.searchParams.set('startPage', String(pageIdx));
  u.searchParams.set('pageSize', String(PAGE_SIZE));
  return u.toString();
}

// Pull the DOI + title of every entry on the current page.
async function extractEntries(page: Page): Promise<{ doi: string; name: string }[]> {
  // NOTE: keep this callback free of named function expressions / `const fn = () =>`
  // helpers. tsx/esbuild's keepNames transform wraps those in a `__name(...)` call
  // that only exists in Node's scope, so it throws "__name is not defined" once the
  // callback is serialized and run inside the browser page.
  return page.$$eval('li.search__item', els =>
    els.map(el => {
      // Collect a DOI reference (a full doi.org URL or an on-site /doi/... path),
      // OR the title link's `id` which is already a bare DOI.
      let idDoi = '';
      let ref   = '';
      // 1) Topic-listing structure (pfirst/otc): explicit DOI link.
      const doiA = el.querySelector('.issue-item__doi a');
      if (doiA) ref = doiA.getAttribute('href') ?? '';
      // 2) Topic-listing title link (id attribute wins, else its /doi/ href).
      if (!ref) {
        const titleA = el.querySelector('.issue-item__title a');
        if (titleA) {
          idDoi = (titleA.getAttribute('id') ?? '').trim();
          if (!idDoi) ref = titleA.getAttribute('href') ?? '';
        }
      }
      // 3) Series/publication structure (aphapts): /doi/book/<doi> title link.
      if (!idDoi && !ref) {
        const metaA = el.querySelector('.meta__title a') ??
                      el.querySelector('.item__body a[href*="/doi/"]');
        if (metaA) ref = metaA.getAttribute('href') ?? '';
      }

      // Normalize the reference to a bare DOI. Handles the /doi/book/, /doi/chapter/,
      // /doi/abs/, etc. prefixes used by the series/publication pages, plus plain
      // /doi/ and full doi.org URLs. Inlined (no helper) — see NOTE above.
      const doi = idDoi ||
        ref.replace(/^https?:\/\/(dx\.)?doi\.org\//, '')
           .replace(/^\/?doi\/(?:book|chapter|abs|full|pdf|epdf)?\/?/, '')
           .trim();

      const titleEl = el.querySelector('.issue-item__title .hlFld-Title') ??
                      el.querySelector('.issue-item__title') ??
                      el.querySelector('.meta__title a') ??
                      el.querySelector('.meta__title');
      const name = (titleEl?.textContent ?? '').trim().replace(/\s+/g, ' ');
      return { doi, name };
    }).filter(x => x.doi)
  );
}

// Does the current page offer a "next page" link (i.e. more pages remain)?
async function hasNextPage(page: Page): Promise<boolean> {
  return (await page.$('nav.pagination a.pagination__btn--next')) !== null;
}

// Read the total-results count shown at the top of the listing (best effort).
async function readTotalCount(page: Page): Promise<number> {
  const txt = await page.$eval('.result__count', el => el.textContent ?? '').catch(() => '');
  const n = parseInt((txt || '').replace(/[^0-9]/g, ''), 10);
  return Number.isFinite(n) ? n : 0;
}

// Walk every page of one source, appending each page's new DOIs to the Books
// tab (and saving) right after that page is scraped — not at the end.
async function scrapeSource(page: Page, source: Source, sink: BookSink): Promise<void> {
  console.log(`\n[Source] ${source.category}  ${source.url}`);
  const seenThisSource = new Set<string>();   // distinct entries walked, to detect the end
  let total = -1;

  for (let pageIdx = 0; pageIdx < MAX_PAGES; pageIdx++) {
    const url = pageUrl(source.url, pageIdx);

    let status = await gotoWithRetry(page, url);
    if (await isCloudflareChallenge(page)) {
      status = await waitForChallengeAndReload(page, url);
      if (await isCloudflareChallenge(page)) {
        console.warn(`  [CF] Challenge not resolved on page ${pageIdx + 1} — stopping this source.`);
        break;
      }
    }
    if (status === 404 || status === 0) {
      console.warn(`  [page ${pageIdx + 1}] status=${status} — stopping this source.`);
      break;
    }

    // Wait for the entry list to render (JS-driven).
    try {
      await page.waitForSelector(ITEM_SELECTOR, { timeout: CONTENT_TIMEOUT });
    } catch {
      console.warn(`  [page ${pageIdx + 1}] no entries rendered — stopping this source.`);
      break;
    }

    if (pageIdx === 0) {
      total = await readTotalCount(page);
      if (total > 0) console.log(`  Listing reports ${total} total entries.`);
    }

    const pageEntries = await extractEntries(page);
    for (const e of pageEntries) seenThisSource.add(normDoi(e.doi));

    // Persist THIS page's changes to the Books tab before moving on.
    const changed = appendEntries(sink, pageEntries, source.category);
    console.log(`  [page ${pageIdx + 1}] ${pageEntries.length} entries (${changed} changed, saved) — ` +
      `walked ${seenThisSource.size}, new rows ${sink.added}, categorized ${sink.updated}`);

    if (pageEntries.length === 0) break;
    if (total > 0 && seenThisSource.size >= total) break;
    if (!(await hasNextPage(page))) break;

    if (DELAY_BETWEEN_PAGES > 0) await page.waitForTimeout(DELAY_BETWEEN_PAGES).catch(() => {});
  }
}

// ---------------------------------------------------------------------------
// Excel read/append
// ---------------------------------------------------------------------------

function normDoi(doi: string): string {
  return doi.trim().toLowerCase();
}

// Zero-based column index of the Category cell (used for in-place updates).
const CATEGORY_COL = COLUMNS.indexOf('Category');

// One row per DOI. `cats` are the categories that DOI has been tagged with; the
// Category cell holds them comma-separated. `row` is the DOI's zero-based sheet
// row so we can update its Category cell in place when a new category is found.
interface DoiRecord {
  row:  number;
  cats: string[];          // display categories, insertion order (no duplicates)
  norm: Set<string>;       // lower-cased categories, for O(1) dedupe
}

// Holds the open workbook plus the running state we mutate as pages are scraped.
// `byDoi` starts with every DOI already in the Books tab (with its existing,
// possibly comma-separated, categories) and grows as we add. A DOI is only ever
// written to one row; seeing it again just merges any new category into that row.
interface BookSink {
  wb: XLSX.WorkBook;
  ws: XLSX.WorkSheet;
  byDoi:   Map<string, DoiRecord>;  // normalized DOI -> its single row + categories
  nextRow: number;        // zero-based index of the next empty row
  added:   number;        // new DOI rows created this run
  updated: number;        // existing DOI rows that gained a new category
  skipped: number;        // entries skipped (DOI already had this category)
}

// Canonical display casing for every known category, keyed by its lower-cased
// form (e.g. "otcdecisiontrees" -> "OtcDecisionTrees"). Built from SOURCES so the
// casing there is the single source of truth.
const CANONICAL_CATS = new Map<string, string>(
  Object.values(SOURCES).map(s => [s.category.toLowerCase(), s.category])
);

// Map a category to its canonical casing if it's a known category; otherwise
// leave it as typed (trimmed). Categories not in SOURCES are passed through.
function canonicalCat(cat: string): string {
  const t = cat.trim();
  return CANONICAL_CATS.get(t.toLowerCase()) ?? t;
}

// Parse a Category cell (possibly "A, B, C") into a clean list of categories,
// normalized to canonical casing and de-duplicated (case-insensitively).
function parseCats(raw: string): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const part of raw.split(',')) {
    const c = canonicalCat(part);
    if (!c) continue;
    const k = c.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(c);
  }
  return out;
}

// Open the workbook and seed `byDoi` from the DOIs already in the Books tab.
// Existing Category cells are rewritten to canonical casing on load; if any cell
// changes the workbook is saved once before scraping begins.
function loadSink(): BookSink {
  const wb = XLSX.readFile(XLSX_FILE);
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found in ${XLSX_FILE}`);
  const rows = XLSX.utils.sheet_to_json<Record<string, string>>(ws, { defval: '' });
  const byDoi = new Map<string, DoiRecord>();
  let recased = 0;
  // rows[i] corresponds to sheet row (i + 1) — row 0 is the header.
  rows.forEach((r, i) => {
    const doi = String(r['DOI'] ?? '').trim();
    if (!doi) return;
    const key  = normDoi(doi);
    const raw  = String(r['Category'] ?? '');
    const cats = parseCats(raw);   // canonical casing, de-duplicated

    // Rewrite this row's own Category cell if canonicalizing changed the text.
    const canonical = cats.join(', ');
    if (canonical !== raw.trim()) { setCell(ws, i + 1, CATEGORY_COL, canonical); recased++; }

    const existing = byDoi.get(key);
    if (existing) {
      // Duplicate DOI row left over from before: merge its categories into the
      // first row we saw and leave this stray row in place (casing already fixed).
      for (const c of cats) {
        if (!existing.norm.has(c.toLowerCase())) {
          existing.norm.add(c.toLowerCase());
          existing.cats.push(c);
        }
      }
    } else {
      byDoi.set(key, { row: i + 1, cats, norm: new Set(cats.map(c => c.toLowerCase())) });
    }
  });

  if (recased > 0) {
    console.log(`[casing] Normalized ${recased} existing Category cell(s) to canonical casing.`);
    try {
      XLSX.writeFile(wb, XLSX_FILE);
    } catch (err) {
      console.warn(`  [xlsx] Could not save casing fixes (${(err as Error).message}). ` +
        `They'll be written on the first page save instead.`);
    }
  }

  // Append below the last used row (header is row 1, data starts row 2).
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');
  return { wb, ws, byDoi, nextRow: range.e.r + 1, added: 0, updated: 0, skipped: 0 };
}

// Set a single cell's string value in place.
function setCell(ws: XLSX.WorkSheet, row: number, col: number, value: string): void {
  ws[XLSX.utils.encode_cell({ r: row, c: col })] = { t: 's', v: value };
}

// Fold one page's entries into the Books tab and SAVE the workbook immediately —
// so progress is persisted after every page, not just at the end of the sweep.
// A brand-new DOI gets a new row (Status/Date blank); a DOI already present just
// gains this category in its existing comma-separated Category cell. Returns the
// number of rows changed (added + updated) this page.
function appendEntries(sink: BookSink, pageEntries: { doi: string; name: string }[], category: string): number {
  const catNorm = category.trim().toLowerCase();
  let changed = 0;
  for (const e of pageEntries) {
    const key = normDoi(e.doi);
    const rec = sink.byDoi.get(key);
    if (rec) {
      if (rec.norm.has(catNorm)) { sink.skipped++; continue; }
      rec.norm.add(catNorm);
      rec.cats.push(category);
      setCell(sink.ws, rec.row, CATEGORY_COL, rec.cats.join(', '));
      sink.updated++;
      changed++;
    } else {
      const row = sink.nextRow;
      const aoaRow = COLUMNS.map(col =>
        col === 'DOI'      ? e.doi :
        col === 'Name'     ? e.name :
        col === 'Category' ? category : '');
      XLSX.utils.sheet_add_aoa(sink.ws, [aoaRow], { origin: row });
      sink.byDoi.set(key, { row, cats: [category], norm: new Set([catNorm]) });
      sink.nextRow++;
      sink.added++;
      changed++;
    }
  }
  if (changed === 0) return 0;

  try {
    XLSX.writeFile(sink.wb, XLSX_FILE);
  } catch (err) {
    // Changes are already in the in-memory sheet, so the next successful save
    // still persists them. Most common cause: the file is open in Excel (lock).
    console.warn(`  [xlsx] Could not save (${(err as Error).message}). ` +
      `Close the file in Excel — these changes will be written on the next save.`);
  }
  return changed;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  console.log('='.repeat(60));
  console.log('Book DOI Scraper — topic listings → Books tab');
  console.log('='.repeat(60));

  const sources = resolveSources();
  console.log(`\nSources (${sources.length}):`);
  for (const s of sources) console.log(`  ${s.category.padEnd(8)} ${s.url}`);

  // Open the workbook and seed the dedupe map from the DOIs already in the tab.
  const sink = loadSink();
  const startingCount = sink.byDoi.size;
  console.log(`\nBooks tab currently has ${startingCount} distinct DOIs.`);

  // ---------------------------------------------------------------------------
  // Browser setup (launch or CDP, mirrors the sibling scripts)
  // ---------------------------------------------------------------------------

  const isHeadless  = process.env.HEADLESS !== 'false';
  const browserMode = (process.env.BROWSER_MODE ?? 'launch').trim().toLowerCase();

  let browser: Browser | null = null;
  let context: BrowserContext;

  if (browserMode === 'cdp') {
    const port = parseInt(process.env.CDP_BASE_PORT ?? '9222', 10);
    console.log(`\n[Browser] mode=cdp  port=${port}`);
    const cdpBrowser = await chromium.connectOverCDP(`http://localhost:${port}`);
    const contexts = cdpBrowser.contexts();
    if (contexts.length === 0) throw new Error(`No browser contexts on port ${port} — run: open-edge.bat`);
    context = contexts[0];
  } else {
    console.log(`\n[Browser] mode=launch  headless=${isHeadless}`);
    browser = await chromium.launch({ headless: isHeadless });
    context = await browser.newContext({ ignoreHTTPSErrors: true });
    await applyAuth(context);
  }

  const page = await context.newPage();

  // ---------------------------------------------------------------------------
  // Scrape every source. Each page's new DOIs are appended + saved as we go
  // (see appendEntries), so the workbook is kept up to date after every page.
  // ---------------------------------------------------------------------------

  try {
    for (const source of sources) {
      await scrapeSource(page, source, sink);
    }
  } finally {
    await page.close().catch(() => {});
    if (browserMode !== 'cdp' && browser) await browser.close().catch(() => {});
  }

  console.log('\n' + '='.repeat(60));
  console.log('Summary:');
  console.log(`  Distinct DOIs at start:          ${startingCount}`);
  console.log(`  New DOI rows added this run:      ${sink.added}`);
  console.log(`  Existing DOIs given a category:  ${sink.updated}`);
  console.log(`  Skipped (DOI already had cat):   ${sink.skipped}`);
  console.log(`  Distinct DOIs now:               ${sink.byDoi.size}`);
  console.log(`\nWorkbook: ${XLSX_FILE}`);
  if (sink.added === 0 && sink.updated === 0) {
    console.log('Nothing changed — Books tab was already up to date.');
  }
}

main().catch(err => {
  console.error('[Fatal]', err);
  process.exit(1);
});
