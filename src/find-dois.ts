import * as dotenv from 'dotenv';
import * as fs     from 'fs';
import * as path   from 'path';
import { chromium, BrowserContext, Browser, Page } from 'playwright';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const ROOT         = path.resolve(__dirname, '..');
const INPUT_CSV    = path.join(ROOT, 'Data', 'Questions DOI extract', 'mapped-question-doi.csv');
const PROD_DOI_CSV = path.join(ROOT, 'Data', 'Questions DOI extract',
  'NAPLEX only production_reporting dim_item 2026-05-27T1520.csv');
const OUTPUT_CSV   = path.join(ROOT, 'output', 'may-2026-review-questions-with-dois.csv');

const BASE_URL            = 'https://pharmacylibrary.com';
const DOI_PREFIX          = '10.21019';
const PAGE_TIMEOUT        = 60_000;
const DELAY_BETWEEN_PAGES = 1_000;
const LOG_THRESHOLD       = 0.50;
const CONCURRENCY         = parseInt(process.env.CONCURRENCY ?? '20', 10);
const BROWSERS            = parseInt(process.env.BROWSERS    ?? '1',  10);

const CHECKPOINT_FILE  = path.join(ROOT, 'output', 'doi-scan-checkpoint.json');
const CHECKPOINT_EVERY = 50;

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
// Fuzzy similarity — Dice coefficient on word sets
// ---------------------------------------------------------------------------

function normalise(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
}

function diceSimilarity(a: string, b: string): number {
  const wa = normalise(a).split(' ').filter(w => w.length > 2);
  const wb = normalise(b).split(' ').filter(w => w.length > 2);
  if (!wa.length || !wb.length) return 0;
  const sa = new Set(wa);
  const sb = new Set(wb);
  const inter = [...sa].filter(w => sb.has(w)).length;
  return (2 * inter) / (sa.size + sb.size);
}

// ---------------------------------------------------------------------------
// Auth — copied from qa-check.ts
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
    console.warn('[Auth] TIP: Set HEADLESS=false, or switch to cookie auth (see .env).');
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
  const username  = process.env.AUTH_USERNAME ?? '';
  const password  = process.env.AUTH_PASSWORD ?? '';

  if (!username || !password) {
    console.warn('[Auth] WARNING: AUTH_USERNAME or AUTH_PASSWORD is empty — pages may return 403');
    return;
  }

  console.log(`[Auth] Navigating to login page: ${loginUrl}`);
  const loginPage = await context.newPage();
  try {
    await loginPage.goto(loginUrl, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT });
    await handleCloudflareChallenge(loginPage);
    if (!loginPage.url().includes('pharmacylibrary.com/action/login') &&
        !loginPage.url().includes('pharmacylibrary.com/login')) {
      await loginPage.goto('https://pharmacylibrary.com/action/login', {
        waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT,
      });
    }
    await loginPage.waitForSelector(
      'input[name="login"], input[name="email"], input[type="email"]',
      { timeout: 10000 }
    ).catch(() => {
      console.warn('[Auth] Login form not found — check AUTH_LOGIN_URL or use HEADLESS=false');
    });
    const usernameField = await loginPage.$('input[name="login"]') ??
                          await loginPage.$('input[name="email"]') ??
                          await loginPage.$('input[type="email"]');
    if (usernameField) await usernameField.fill(username);
    const passwordField = await loginPage.$('input[name="password"]') ??
                          await loginPage.$('input[type="password"]');
    if (passwordField) await passwordField.fill(password);
    await loginPage.click('button[type="submit"], input[type="submit"]');
    await loginPage.waitForNavigation({ waitUntil: 'networkidle', timeout: PAGE_TIMEOUT })
      .catch(() => {});
    const finalUrl = loginPage.url();
    if (finalUrl.includes('/login') || finalUrl.includes('/signin')) {
      console.warn('[Auth] WARNING: Still on login page — credentials may be wrong or Turnstile blocked');
      console.warn('[Auth] TIP: Try AUTH_METHOD=cookie in .env');
    } else {
      console.log(`[Auth] Login succeeded → ${finalUrl}`);
    }
  } finally {
    await loginPage.close();
  }
}

// ---------------------------------------------------------------------------
// Cloudflare challenge helpers — copied from qa-check.ts
// ---------------------------------------------------------------------------

async function isCloudflareChallenge(page: Page): Promise<boolean> {
  try {
    const title = await page.title();
    if (
      title.toLowerCase().includes('security verification') ||
      title.toLowerCase().includes('just a moment') ||
      title.toLowerCase().includes('attention required')
    ) return true;
    const bodySnippet = await page.evaluate(
      () => (document.body?.innerText ?? '').slice(0, 500)
    );
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
  ).catch(() => { console.warn('[CF] Challenge did not resolve in 30 s — retrying navigation.'); });
  await page.waitForTimeout(1500);
  const response = await page.goto(url, {
    waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT,
  }).catch(() => null);
  return response?.status() ?? 0;
}

// ---------------------------------------------------------------------------
// Checkpoint helpers
// ---------------------------------------------------------------------------

interface Checkpoint {
  visitedDois: string[];
  candidates:  Candidate[];
}

function loadCheckpoint(): Checkpoint {
  if (!fs.existsSync(CHECKPOINT_FILE)) return { visitedDois: [], candidates: [] };
  try {
    return JSON.parse(fs.readFileSync(CHECKPOINT_FILE, 'utf-8')) as Checkpoint;
  } catch {
    console.warn('[Checkpoint] Could not parse checkpoint file — starting fresh.');
    return { visitedDois: [], candidates: [] };
  }
}

function saveCheckpoint(visited: Set<string>, candidates: Candidate[]): void {
  fs.mkdirSync(path.dirname(CHECKPOINT_FILE), { recursive: true });
  fs.writeFileSync(
    CHECKPOINT_FILE,
    JSON.stringify({ visitedDois: [...visited], candidates }, null, 2),
    'utf-8'
  );
}

// ---------------------------------------------------------------------------
// Navigation with retry
// ---------------------------------------------------------------------------

async function gotoWithRetry(
  page: Page,
  url: string,
  maxAttempts = 3,
  baseDelayMs = 2_000,
): Promise<number> {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp   = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT });
      const status = resp?.status() ?? 0;
      if (status === 404) return 404;
      if (status >= 200 && status < 300) return status;
      if (attempt < maxAttempts) {
        printWarning(`  [Retry][${attempt}/${maxAttempts}] status=${status} — waiting ${baseDelayMs * attempt}ms`);
        await page.waitForTimeout(baseDelayMs * attempt);
      } else {
        return status;
      }
    } catch (err: unknown) {
      const msg = (err as Error).message ?? '';
      if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) throw err;
      if (attempt < maxAttempts) {
        printWarning(`  [Retry][${attempt}/${maxAttempts}] nav error: ${msg.slice(0, 80)}`);
        await page.waitForTimeout(baseDelayMs * attempt);
      } else {
        return 0;
      }
    }
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Progress logging — single updating line via \r
// ---------------------------------------------------------------------------

let _progressLine = '';

function logProgress(visited: number, total: number, candidateCount: number, startTime: number): void {
  const pct     = ((visited / total) * 100).toFixed(1);
  const elapsed = Math.round((Date.now() - startTime) / 1000);
  _progressLine = `[Progress] ${visited}/${total} (${pct}%) | candidates: ${candidateCount} | elapsed: ${elapsed}s`;
  process.stdout.write('\r' + _progressLine.padEnd(110));
}

function printMatch(line: string): void {
  process.stdout.write('\r\x1B[K');
  process.stdout.write(line + '\n');
  if (_progressLine) process.stdout.write(_progressLine.padEnd(110));
}

function printWarning(line: string): void {
  process.stdout.write('\r\x1B[K');
  process.stdout.write(line + '\n');
  if (_progressLine) process.stdout.write(_progressLine.padEnd(110));
}

// ---------------------------------------------------------------------------
// Browser slot — wraps one browser instance and its liveness flag
// ---------------------------------------------------------------------------

interface BrowserSlot {
  browser: Browser | null;  // null for CDP mode
  context: BrowserContext;
  alive:   boolean;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

interface ReviewQuestion {
  question_no:   string;
  question_stem: string;
  answer:        string;
  doi:           string;
  match_score:   string;
}

interface Candidate {
  question_no:   string;
  question_stem: string;
  doi:           string;
  match_score:   string;
}

async function main(): Promise<void> {
  console.log('='.repeat(60));
  console.log('DOI Finder — live URL matching');
  console.log('='.repeat(60));

  // Load review questions
  const { rows } = parseSimpleCsv(fs.readFileSync(INPUT_CSV, 'utf-8'));
  const allQuestions: ReviewQuestion[] = rows.map(r => ({
    question_no:   r['question_no']   ?? '',
    question_stem: r['question_stem'] ?? '',
    answer:        r['answer']        ?? '',
    doi:           r['doi']           ?? '',
    match_score:   '',
  }));

  // Only match questions that don't already have a DOI in the input CSV
  const questions = allQuestions.filter(q => !q.doi.trim());
  console.log(`\nLoaded ${allQuestions.length} questions from review CSV (${questions.length} without a DOI)`);

  // Collect DOI suffixes already assigned in the input CSV so we can skip them
  const assignedDoiSuffixes = new Set<string>(
    allQuestions
      .map(q => { const d = q.doi.trim(); const i = d.indexOf('/'); return i >= 0 ? d.slice(i + 1) : d; })
      .filter(Boolean)
  );

  // Load production DOI list — header is "Item Item DOI", each row is a full DOI
  const { rows: doiRows } = parseSimpleCsv(fs.readFileSync(PROD_DOI_CSV, 'utf-8'));
  const doiSuffixes: string[] = doiRows
    .map(r => {
      const full = r['Item Item DOI'] ?? '';
      const slash = full.indexOf('/');
      return slash >= 0 ? full.slice(slash + 1) : '';
    })
    .filter(s => s.length > 0 && !assignedDoiSuffixes.has(s));
  console.log(`Loaded ${doiSuffixes.length} DOIs from production CSV (excluding ${assignedDoiSuffixes.size} already assigned)`);

  if (questions.length === 0) {
    console.log('\nAll questions already have DOIs assigned — nothing to scan.');
    return;
  }

  // All candidate matches found across all DOI pages
  const candidates: Candidate[] = [];

  // Load checkpoint (resumes a previous partial run)
  const checkpoint  = loadCheckpoint();
  const visitedDois = new Set<string>(checkpoint.visitedDois);
  candidates.push(...checkpoint.candidates);
  if (visitedDois.size > 0) {
    console.log(`[Checkpoint] Resuming — ${visitedDois.size} DOIs already visited, ${candidates.length} candidates loaded.`);
  }

  const remainingDois = doiSuffixes.filter(s => !visitedDois.has(s));
  let nextDoiIdx = 0;
  let processedSinceCheckpoint = 0;
  let totalVisited = visitedDois.size;
  const startTime  = Date.now();

  // ---------------------------------------------------------------------------
  // Browser setup
  // ---------------------------------------------------------------------------

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
      br.on('disconnected', () => {
        slot.alive = false;
        printWarning(`[Browser][Slot${b + 1}] Disconnected — its workers will stop.`);
      });
      ctx.on('close', () => { slot.alive = false; });
      slots.push(slot);
    }
  }

  // ---------------------------------------------------------------------------
  // Worker
  // ---------------------------------------------------------------------------

  const CONTENT_SELECTOR = '.do .question, #question, .qna-statement';
  const CONTENT_TIMEOUT  = 15_000;

  async function runWorker(workerPage: Page, workerId: number, slot: BrowserSlot): Promise<void> {
    while (slot.alive) {
      const i = nextDoiIdx++;
      if (i >= remainingDois.length) break;

      const doiSuffix = remainingDois[i];
      const url = `${BASE_URL}/do/${DOI_PREFIX}/${doiSuffix}/abs/`;

      try {
        let status: number;
        try {
          status = await gotoWithRetry(workerPage, url);
        } catch (err: unknown) {
          const msg = (err as Error).message ?? String(err);
          if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) {
            slot.alive = false;
            break;
          }
          visitedDois.add(doiSuffix);
          continue;
        }

        if (await isCloudflareChallenge(workerPage)) {
          status = await waitForChallengeAndReload(workerPage, url);
          if (await isCloudflareChallenge(workerPage)) {
            printWarning(`  [CF][W${workerId}] Challenge not resolved for ${doiSuffix} — skipping`);
            visitedDois.add(doiSuffix);
            continue;
          }
        }

        if (status === 404 || status === 0) {
          visitedDois.add(doiSuffix);
          continue;
        }

        // Wait for JS-rendered question content, with one re-navigation retry on timeout
        let containerFound = false;
        try {
          await workerPage.waitForSelector(CONTENT_SELECTOR, { timeout: CONTENT_TIMEOUT });
          containerFound = true;
        } catch {
          printWarning(`  [W${workerId}] Content timeout for ${doiSuffix} — retrying nav...`);
          try {
            await gotoWithRetry(workerPage, url, 2, 3_000);
            await workerPage.waitForSelector(CONTENT_SELECTOR, { timeout: CONTENT_TIMEOUT });
            containerFound = true;
          } catch {
            printWarning(`  [W${workerId}] Content not found after retry for ${doiSuffix} — skipping`);
          }
        }

        if (!containerFound) { visitedDois.add(doiSuffix); continue; }

        const container = await workerPage.$('.do .question') ??
                          await workerPage.$('#question') ??
                          await workerPage.$('.qna-statement');
        if (!container) { visitedDois.add(doiSuffix); continue; }
        const paras = await container.$$('p');
        if (paras.length === 0) { visitedDois.add(doiSuffix); continue; }
        const texts = await Promise.all(paras.map(p => p.innerText()));
        const pageStem = texts.map(t => t.trim()).filter(Boolean).join(' ');
        if (!pageStem) { visitedDois.add(doiSuffix); continue; }

        // Score against every unresolved question
        for (let idx = 0; idx < questions.length; idx++) {
          const score = diceSimilarity(pageStem, questions[idx].question_stem);
          if (score >= LOG_THRESHOLD) {
            const q = questions[idx];
            const scoreStr = (score * 100).toFixed(1) + '%';
            const doi = `${DOI_PREFIX}/${doiSuffix}`;
            candidates.push({ question_no: q.question_no, question_stem: pageStem, doi, match_score: scoreStr });
            printMatch(
              `  Q${q.question_no.padStart(3)}  ${doiSuffix.padEnd(22)} ${scoreStr.padStart(6)}  ${pageStem.slice(0, 60)}`
            );
          }
        }

        visitedDois.add(doiSuffix);
        totalVisited++;
        processedSinceCheckpoint++;

        if (processedSinceCheckpoint >= CHECKPOINT_EVERY) {
          saveCheckpoint(visitedDois, candidates);
          processedSinceCheckpoint = 0;
        }

        logProgress(totalVisited, doiSuffixes.length, candidates.length, startTime);

      } catch (err: unknown) {
        const msg = (err as Error).message ?? String(err);
        if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) {
          printWarning(`[W${workerId}][Stopped] Browser closed — stopping worker.`);
          slot.alive = false;
          break;
        }
      }

      if (slot.alive) {
        await workerPage.waitForTimeout(DELAY_BETWEEN_PAGES).catch(() => {});
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Create pages round-robin across slots and launch workers
  // ---------------------------------------------------------------------------

  const workerEntries: { page: Page; slot: BrowserSlot }[] = [];
  for (let i = 0; i < CONCURRENCY; i++) {
    const slot = slots[i % slots.length];
    const page = await slot.context.newPage();
    workerEntries.push({ page, slot });
  }

  console.log(`\nVisiting ${remainingDois.length} remaining DOIs (${visitedDois.size} already done) across ${CONCURRENCY} parallel tabs (${slots.length} browser(s))...\n`);
  console.log(`  ${'Q#'.padEnd(5)} ${'DOI suffix'.padEnd(22)} ${'Score'.padEnd(7)} Stem preview`);
  console.log('  ' + '-'.repeat(80));

  await Promise.all(workerEntries.map(({ page, slot }, i) => runWorker(page, i + 1, slot)));

  // ---------------------------------------------------------------------------
  // Teardown
  // ---------------------------------------------------------------------------

  await Promise.all(
    workerEntries.map(({ page, slot }) =>
      slot.alive ? page.close().catch(() => {}) : Promise.resolve()
    )
  );

  if (browserMode !== 'cdp') {
    const closed = new Set<BrowserSlot>();
    for (const slot of slots) {
      if (!closed.has(slot) && slot.browser) {
        closed.add(slot);
        await slot.browser.close().catch(() => {});
      }
    }
  }

  // Final checkpoint save
  saveCheckpoint(visitedDois, candidates);
  console.log('\n[Checkpoint] Final checkpoint saved.');
  if (visitedDois.size >= doiSuffixes.length) {
    console.log('[Checkpoint] All DOIs scanned. Delete output/doi-scan-checkpoint.json before running a fresh full scan.');
  }

  // Sort candidates: by question_no numerically, then match_score descending
  candidates.sort((a, b) => {
    const nDiff = parseInt(a.question_no) - parseInt(b.question_no);
    if (nDiff !== 0) return nDiff;
    return parseFloat(b.match_score) - parseFloat(a.match_score);
  });

  // Merge with existing output (if any) so successive runs accumulate candidates
  fs.mkdirSync(path.dirname(OUTPUT_CSV), { recursive: true });
  const OUT_HEADERS = ['question_no', 'question_stem', 'doi', 'match_score'];
  if (fs.existsSync(OUTPUT_CSV)) {
    const existing = parseSimpleCsv(fs.readFileSync(OUTPUT_CSV, 'utf-8'));
    for (const row of existing.rows) {
      const alreadyPresent = candidates.some(
        c => c.question_no === row['question_no'] && c.doi === row['doi']
      );
      if (!alreadyPresent) {
        candidates.push({
          question_no:   row['question_no']   ?? '',
          question_stem: row['question_stem'] ?? '',
          doi:           row['doi']           ?? '',
          match_score:   row['match_score']   ?? '',
        });
      }
    }
    candidates.sort((a, b) => {
      const nDiff = parseInt(a.question_no) - parseInt(b.question_no);
      if (nDiff !== 0) return nDiff;
      return parseFloat(b.match_score) - parseFloat(a.match_score);
    });
  }
  const outLines = [OUT_HEADERS.join(',')];
  for (const c of candidates) {
    outLines.push(OUT_HEADERS.map(h => escapeCsv((c as unknown as Record<string, string>)[h] ?? '')).join(','));
  }
  fs.writeFileSync(OUTPUT_CSV, outLines.join('\n') + '\n', 'utf-8');

  const questionNosWithCandidates = new Set(candidates.map(c => c.question_no));
  console.log('\n' + '='.repeat(60));
  console.log('Results:');
  console.log(`  Total candidates logged (>= ${(LOG_THRESHOLD * 100).toFixed(0)}%): ${candidates.length}`);
  console.log(`  Questions with at least one candidate: ${questionNosWithCandidates.size} / ${questions.length}`);
  console.log(`\nOutput: ${OUTPUT_CSV}`);
}

main().catch(err => {
  console.error('[Fatal]', err);
  process.exit(1);
});
