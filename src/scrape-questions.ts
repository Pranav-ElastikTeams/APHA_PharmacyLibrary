import * as dotenv from 'dotenv';
import * as fs     from 'fs';
import * as path   from 'path';
import { chromium, BrowserContext, Browser, Page } from 'playwright';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const ROOT         = path.resolve(__dirname, '..');
const PROD_DOI_CSV = path.join(ROOT, 'Data', 'Questions DOI extract',
  'NAPLEX only production_reporting dim_item 2026-05-27T1520.csv');

const DOI_HEADER      = 'Item Item DOI';   // first column of the production CSV (left untouched)
const QUESTION_HEADER = 'question_text';   // second column we add / update

// Where to WRITE the augmented CSV. By default we write the question_text column
// straight back into the production CSV (in place — DOIs are only READ, never
// added/removed/reordered). Set SCRAPE_OUTPUT to write a separate copy instead.
function resolveTarget(): string {
  const v = process.env.SCRAPE_OUTPUT;
  if (v && v.trim()) return path.isAbsolute(v) ? v : path.join(ROOT, v);
  return PROD_DOI_CSV;
}

const BASE_URL            = 'https://pharmacylibrary.com';
const DOI_PREFIX          = '10.21019';
const PAGE_TIMEOUT        = 60_000;
const DELAY_BETWEEN_PAGES = parseInt(process.env.SCRAPE_DELAY_MS ?? '500', 10);
const CONCURRENCY         = parseInt(process.env.CONCURRENCY ?? '20', 10);
const BROWSERS            = parseInt(process.env.BROWSERS    ?? '1',  10);

const CONTENT_SELECTOR = '.do .question, #question, .qna-statement';
const CONTENT_TIMEOUT  = parseInt(process.env.SCRAPE_CONTENT_TIMEOUT ?? '12000', 10);

const CHECKPOINT_FILE  = path.join(ROOT, 'output', 'scrape-questions-checkpoint.json');
const CHECKPOINT_EVERY = 50;

// How often (in DOIs completed THIS run) to refresh the progress line, and what
// it counts against. Because a run only works the still-empty question_text rows,
// progress is measured over those remaining DOIs — not the whole CSV.
//   SCRAPE_PROGRESS_EVERY = 1  → update on every DOI (default)
//   SCRAPE_PROGRESS_EVERY = N  → update every N DOIs (quieter logs on big runs)
//   SCRAPE_PROGRESS_EVERY = 0  → disable the progress line entirely
const PROGRESS_EVERY = parseInt(process.env.SCRAPE_PROGRESS_EVERY ?? '1', 10);

// --- Dynamic test setup (read from .env) -----------------------------------
// SCRAPE_DOIS  : comma-separated DOIs (full "10.21019/qna-XXXXXX" or bare
//                "qna-XXXXXX"). When set, the production CSV is IGNORED and
//                ONLY these DOIs are scraped — handy for a quick smoke test.
// SCRAPE_LIMIT : cap the number of DOIs processed this run (0 = all).
// SCRAPE_FRESH : "true" → ignore any existing checkpoint and start over.
const TEST_DOIS  = (process.env.SCRAPE_DOIS ?? '')
  .split(',')
  .map(s => s.trim())
  .filter(Boolean)
  .map(s => { const i = s.indexOf('/'); return i >= 0 ? s.slice(i + 1) : s; });
const SCRAPE_LIMIT = parseInt(process.env.SCRAPE_LIMIT ?? '0', 10);
const SCRAPE_FRESH = (process.env.SCRAPE_FRESH ?? 'false').trim().toLowerCase() === 'true';

// ---------------------------------------------------------------------------
// CSV helpers
// ---------------------------------------------------------------------------

function parseCsvLine(line: string): string[] {
  const out: string[] = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === ',' && !inQ) {
      out.push(cur); cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function parseSimpleCsv(content: string): { headers: string[]; rows: Record<string, string>[] } {
  const lines = content.split(/\r?\n/);
  const headers = parseCsvLine(lines[0]);
  const rows: Record<string, string>[] = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const vals = parseCsvLine(line);
    const obj: Record<string, string> = {};
    headers.forEach((h, idx) => { obj[h] = vals[idx] ?? ''; });
    rows.push(obj);
  }
  return { headers, rows };
}

function escapeCsv(v: string): string {
  if (v.includes(',') || v.includes('"') || v.includes('\n'))
    return `"${v.replace(/"/g, '""')}"`;
  return v;
}

// ---------------------------------------------------------------------------
// Auth (cookie/login) — copied from find-dois.ts, only used in launch mode
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
  printWarning('[CF] Challenge detected — waiting up to 30 s for auto-resolve...');
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
// Progress logging
// ---------------------------------------------------------------------------

let _progressLine = '';

function logProgress(done: number, total: number, found: number, startTime: number): void {
  const pct     = ((done / total) * 100).toFixed(1);
  const elapsed = Math.round((Date.now() - startTime) / 1000);
  _progressLine = `[Progress] ${done}/${total} (${pct}%) | with text: ${found} | elapsed: ${elapsed}s`;
  process.stdout.write('\r' + _progressLine.padEnd(110));
}

function printLine(line: string): void {
  process.stdout.write('\r\x1B[K');
  process.stdout.write(line + '\n');
  if (_progressLine) process.stdout.write(_progressLine.padEnd(110));
}
const printWarning = printLine;

// ---------------------------------------------------------------------------
// Checkpoint
// ---------------------------------------------------------------------------

interface Checkpoint {
  results: [string, string][];   // [doiSuffix, questionText] — text '' means visited-but-empty
}

function loadCheckpoint(): Map<string, string> {
  if (SCRAPE_FRESH || !fs.existsSync(CHECKPOINT_FILE)) return new Map();
  try {
    const cp = JSON.parse(fs.readFileSync(CHECKPOINT_FILE, 'utf-8')) as Checkpoint;
    return new Map(cp.results);
  } catch {
    console.warn('[Checkpoint] Could not parse checkpoint — starting fresh.');
    return new Map();
  }
}

function saveCheckpoint(results: Map<string, string>): void {
  fs.mkdirSync(path.dirname(CHECKPOINT_FILE), { recursive: true });
  fs.writeFileSync(CHECKPOINT_FILE, JSON.stringify({ results: [...results] }), 'utf-8');
}

// Augment the production CSV: keep the DOI column and row order EXACTLY as-is,
// and write each row's scraped question into a second `question_text` column.
// Rows not yet scraped keep whatever question text is already on them (empty
// on the first run). The production CSV's line ending is preserved.
function writeOutput(results: Map<string, string>): void {
  const target = resolveTarget();
  const raw = fs.readFileSync(PROD_DOI_CSV, 'utf-8');   // production CSV is the row template
  const eol = raw.includes('\r\n') ? '\r\n' : '\n';
  const lines = raw.split(/\r?\n/);
  while (lines.length && lines[lines.length - 1] === '') lines.pop();
  if (lines.length === 0) return;

  const header = parseCsvLine(lines[0]);
  const out: string[] = [`${escapeCsv(header[0] || DOI_HEADER)},${QUESTION_HEADER}`];
  for (let i = 1; i < lines.length; i++) {
    const fields  = parseCsvLine(lines[i]);
    const doiFull = fields[0] ?? '';
    if (!doiFull) continue;
    const slash   = doiFull.indexOf('/');
    const suffix  = slash >= 0 ? doiFull.slice(slash + 1) : doiFull;
    const existing = fields[1] ?? '';
    const text = results.has(suffix) ? (results.get(suffix) as string) : existing;
    out.push(`${escapeCsv(doiFull)},${escapeCsv(text)}`);
  }

  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, out.join(eol) + eol, 'utf-8');
}

// ---------------------------------------------------------------------------
// Browser slot
// ---------------------------------------------------------------------------

interface BrowserSlot {
  browser: Browser | null;
  context: BrowserContext;
  alive:   boolean;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  console.log('='.repeat(60));
  console.log('Question Scraper — DOI → question text');
  console.log('='.repeat(60));

  // Build the DOI list: explicit test list, or the production CSV.
  // Either way the production CSV is only READ — DOIs are never added/removed.
  const prodParsed = parseSimpleCsv(fs.readFileSync(PROD_DOI_CSV, 'utf-8'));
  let doiSuffixes: string[];
  if (TEST_DOIS.length > 0) {
    doiSuffixes = TEST_DOIS;
    console.log(`\n[Test mode] SCRAPE_DOIS set — scraping ${doiSuffixes.length} explicit DOI(s), ignoring production CSV.`);
  } else {
    doiSuffixes = prodParsed.rows
      .map(r => { const full = r[prodParsed.headers[0]] ?? ''; const i = full.indexOf('/'); return i >= 0 ? full.slice(i + 1) : ''; })
      .filter(Boolean);
    console.log(`\nLoaded ${doiSuffixes.length} DOIs from production CSV.`);
  }

  // Resume from checkpoint: skip DOIs already scraped.
  const results = loadCheckpoint();
  if (results.size > 0) {
    console.log(`[Checkpoint] Resuming — ${results.size} DOIs already scraped (use SCRAPE_FRESH=true to restart).`);
  }

  // Also seed from any question_text already saved in the production CSV's second
  // column, so a re-run skips already-filled rows (unless SCRAPE_FRESH=true).
  if (!SCRAPE_FRESH && TEST_DOIS.length === 0 && prodParsed.headers.length > 1) {
    let seeded = 0;
    for (const r of prodParsed.rows) {
      const full = r[prodParsed.headers[0]] ?? '';
      const i = full.indexOf('/'); const suffix = i >= 0 ? full.slice(i + 1) : full;
      const text = (r[prodParsed.headers[1]] ?? '').trim();
      if (suffix && text && !results.has(suffix)) { results.set(suffix, text); seeded++; }
    }
    if (seeded > 0) console.log(`[Resume] ${seeded} rows already have question_text in the CSV — skipping them.`);
  }

  // Only scrape DOIs whose question_text is still EMPTY. Rows that already have
  // text (seeded from the CSV's second column above, or a non-empty checkpoint
  // entry) are skipped — we never re-scrape / overwrite a filled question_text.
  // Note: an empty checkpoint entry (value '') does NOT count as "done", so a
  // previously-empty DOI is retried on the next run.
  let remaining = doiSuffixes.filter(s => !(results.get(s) ?? '').trim());
  if (SCRAPE_LIMIT > 0 && remaining.length > SCRAPE_LIMIT) {
    remaining = remaining.slice(0, SCRAPE_LIMIT);
    console.log(`[Test mode] SCRAPE_LIMIT=${SCRAPE_LIMIT} — capping this run to ${remaining.length} DOIs.`);
  }

  if (remaining.length === 0) {
    console.log('\nNothing to scrape — all requested DOIs already done. Writing output and exiting.');
    writeOutput(results);
    console.log(`Output: ${resolveTarget()}`);
    return;
  }

  let nextIdx = 0;
  let processedSinceCheckpoint = 0;
  let processedThisRun = 0;            // DOIs completed in THIS run (progress denominator basis)
  let withText = 0;                    // of those, how many yielded question text
  const startTime = Date.now();
  const totalTarget = remaining.length;

  // -------------------------------------------------------------------------
  // Browser setup
  // -------------------------------------------------------------------------

  const isHeadless  = process.env.HEADLESS !== 'false';
  const browserMode = (process.env.BROWSER_MODE ?? 'launch').trim().toLowerCase();
  let slots: BrowserSlot[];

  if (browserMode === 'cdp') {
    const browserCount = Math.max(1, BROWSERS);
    const basePort     = parseInt(process.env.CDP_BASE_PORT ?? '9222', 10);
    console.log(`\n[Browser] mode=cdp  concurrency=${CONCURRENCY}  browsers=${browserCount}  base-port=${basePort}`);
    slots = [];
    for (let b = 0; b < browserCount; b++) {
      const port   = basePort + b;
      const cdpUrl = `http://localhost:${port}`;
      console.log(`[Browser] Connecting to Edge via CDP: ${cdpUrl}`);
      const cdpBrowser = await chromium.connectOverCDP(cdpUrl);
      const existing = cdpBrowser.contexts();
      if (existing.length === 0) {
        throw new Error(`No browser contexts found on port ${port} — run: open-edge.bat ${browserCount}`);
      }
      const slot: BrowserSlot = { browser: null, context: existing[0], alive: true };
      cdpBrowser.on('disconnected', () => {
        slot.alive = false;
        printWarning(`[Browser][CDP:${port}] Disconnected — its workers will stop.`);
      });
      slot.context.on('close', () => { slot.alive = false; });
      slots.push(slot);
    }
  } else {
    const browserCount = Math.max(1, BROWSERS);
    console.log(`\n[Browser] mode=launch  headless=${isHeadless}  concurrency=${CONCURRENCY}  browsers=${browserCount}`);
    slots = [];
    for (let b = 0; b < browserCount; b++) {
      const br  = await chromium.launch({ headless: isHeadless });
      const ctx = await br.newContext({ ignoreHTTPSErrors: true });
      await applyAuth(ctx);
      const slot: BrowserSlot = { browser: br, context: ctx, alive: true };
      br.on('disconnected', () => { slot.alive = false; });
      ctx.on('close', () => { slot.alive = false; });
      slots.push(slot);
    }
  }

  // -------------------------------------------------------------------------
  // Worker
  // -------------------------------------------------------------------------

  async function runWorker(workerPage: Page, workerId: number, slot: BrowserSlot): Promise<void> {
    while (slot.alive) {
      const i = nextIdx++;
      if (i >= remaining.length) break;

      const doiSuffix = remaining[i];
      const url = `${BASE_URL}/do/${DOI_PREFIX}/${doiSuffix}/abs/`;

      try {
        let status: number;
        try {
          status = await gotoWithRetry(workerPage, url);
        } catch (err: unknown) {
          const msg = (err as Error).message ?? String(err);
          if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) {
            slot.alive = false; break;
          }
          results.set(doiSuffix, '');
          continue;
        }

        if (await isCloudflareChallenge(workerPage)) {
          status = await waitForChallengeAndReload(workerPage, url);
          if (await isCloudflareChallenge(workerPage)) {
            printWarning(`  [CF][W${workerId}] Challenge not resolved for ${doiSuffix} — skipping`);
            results.set(doiSuffix, '');
            continue;
          }
        }

        if (status === 404 || status === 0) {
          results.set(doiSuffix, '');
          continue;
        }

        // Wait for JS-rendered question content, with one re-navigation retry.
        let containerFound = false;
        try {
          await workerPage.waitForSelector(CONTENT_SELECTOR, { timeout: CONTENT_TIMEOUT });
          containerFound = true;
        } catch {
          try {
            await gotoWithRetry(workerPage, url, 2, 3_000);
            await workerPage.waitForSelector(CONTENT_SELECTOR, { timeout: CONTENT_TIMEOUT });
            containerFound = true;
          } catch {
            printWarning(`  [W${workerId}] No question content for ${doiSuffix} — recording empty`);
          }
        }

        let questionText = '';
        if (containerFound) {
          const container = await workerPage.$('.do .question') ??
                            await workerPage.$('#question') ??
                            await workerPage.$('.qna-statement');
          if (container) {
            const paras = await container.$$('p');
            if (paras.length > 0) {
              const texts = await Promise.all(paras.map(p => p.innerText()));
              questionText = texts.map(t => t.trim().replace(/\s+/g, ' ')).filter(Boolean).join(' ');
            } else {
              questionText = (await container.innerText()).trim().replace(/\s+/g, ' ');
            }
          }
        }

        results.set(doiSuffix, questionText);
        if (questionText) {
          withText++;
          printLine(`  ${doiSuffix.padEnd(22)} ${questionText.slice(0, 70)}`);
        }

        processedSinceCheckpoint++;
        if (processedSinceCheckpoint >= CHECKPOINT_EVERY) {
          saveCheckpoint(results);
          writeOutput(results);
          processedSinceCheckpoint = 0;
        }
        processedThisRun++;
        if (PROGRESS_EVERY > 0 &&
            (processedThisRun % PROGRESS_EVERY === 0 || processedThisRun === totalTarget)) {
          logProgress(processedThisRun, totalTarget, withText, startTime);
        }

      } catch (err: unknown) {
        const msg = (err as Error).message ?? String(err);
        if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) {
          printWarning(`[W${workerId}][Stopped] Browser closed — stopping worker.`);
          slot.alive = false; break;
        }
      }

      if (slot.alive && DELAY_BETWEEN_PAGES > 0) {
        await workerPage.waitForTimeout(DELAY_BETWEEN_PAGES).catch(() => {});
      }
    }
  }

  // -------------------------------------------------------------------------
  // Launch workers round-robin across slots
  // -------------------------------------------------------------------------

  const workerEntries: { page: Page; slot: BrowserSlot }[] = [];
  const workerTabs = Math.min(CONCURRENCY, Math.max(1, totalTarget));
  for (let i = 0; i < workerTabs; i++) {
    const slot = slots[i % slots.length];
    const page = await slot.context.newPage();
    workerEntries.push({ page, slot });
  }

  console.log(`\nScraping ${remaining.length} DOIs across ${workerTabs} parallel tab(s) (${slots.length} browser(s))...\n`);
  await Promise.all(workerEntries.map(({ page, slot }, i) => runWorker(page, i + 1, slot)));

  // -------------------------------------------------------------------------
  // Teardown
  // -------------------------------------------------------------------------

  await Promise.all(workerEntries.map(({ page, slot }) =>
    slot.alive ? page.close().catch(() => {}) : Promise.resolve()));

  if (browserMode !== 'cdp') {
    const closed = new Set<BrowserSlot>();
    for (const slot of slots) {
      if (!closed.has(slot) && slot.browser) { closed.add(slot); await slot.browser.close().catch(() => {}); }
    }
  }

  saveCheckpoint(results);
  writeOutput(results);

  const found = [...results.values()].filter(Boolean).length;
  console.log('\n' + '='.repeat(60));
  console.log('Results:');
  console.log(`  DOIs scraped (total in checkpoint): ${results.size}`);
  console.log(`  With question text:                 ${found}`);
  console.log(`  Empty / 404 / no content:           ${results.size - found}`);
  console.log(`\nOutput: ${resolveTarget()}`);
}

main().catch(err => {
  console.error('[Fatal]', err);
  process.exit(1);
});
