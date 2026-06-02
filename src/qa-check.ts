import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { chromium, BrowserContext, Browser, Page, ConsoleMessage } from 'playwright';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const EXCEL_DOI_PATH = path.resolve(__dirname, '..', 'Data', 'NAPLEX review', 'NAPLEX Math Review questions_29-4-26.xlsx');
const EXCEL_QA_PATH  = path.resolve(__dirname, '..', 'Data', 'NAPLEX review', 'Naplex_Math-Review-Questions.xlsx');
const CSV_PATH        = path.resolve(__dirname, '..', 'output', 'qna_qa_report.csv');
const SCREENSHOTS_DIR = path.resolve(__dirname, '..', 'screenshots');

const BASE_URL   = 'https://staging.pharmacylibrary.com';
const DOI_PREFIX = '10.21019';

const PAGE_TIMEOUT        = 60_000;
const DELAY_BETWEEN_PAGES = 1_000;

// ---------------------------------------------------------------------------
// CSS Selectors — confirmed against live page HTML
// ---------------------------------------------------------------------------

const SELECTORS = {
  // Question stem: <div class="qna-statement" id="question"><p>...</p></div>
  questionStem:    '#question p, .qna-statement p',

  // Reveal button: clicking shows #answer-block
  revealButton:    'input[name="choice"]',

  // Answer block wrapper (has class "hide" initially, removed on click)
  answerBlock:     '#answer-block',

  // The displayed correct answer text
  answerText:      '#answer-block span#answer, #answer-block .answer',

  // Explanation
  explanation:     '.qna-explanation',

  // Math equation images (equations rendered as <img> inside .qna-equation spans)
  equationImages:  '.qna-equation img',
};

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface PageResult {
  doi: string;
  question_no: number | null;       // S.No. from DOI file
  url: string;
  http_status: number;
  has_question_text: boolean | null;
  page_stem: string;                // question stem as rendered on the page
  reveal_button_present: boolean | null;
  displayed_answer: string;
  expected_answer: string;
  answer_matches: boolean | null;
  broken_equation_images: string;
  broken_other_images: string;
  raw_mathml_visible: boolean;
  raw_entities_visible: boolean;
  explanation_present: boolean | null;
  missing_alt_images: string;
  console_errors: string;
  other_issues: string;
  notes: string;
  hasIssues: boolean;
}

interface SummaryStats {
  total: number;
  issues: number;
  notFound: number;
  missingQuestionText: number;
  missingRevealButton: number;
  answerMismatches: number;
  brokenEquationImages: number;
  brokenOtherImages: number;
  rawMathml: number;
  rawEntities: number;
  missingExplanation: number;
  consoleErrors: number;
  unicodeReplacement: number;
}

interface QuestionData {
  stem: string;
  answer: string;
  serialNo?: number;
}

// ---------------------------------------------------------------------------
// Excel helpers
// ---------------------------------------------------------------------------

// Merged cells in Excel only store a value in the top-left cell of the merge.
// This fills every cell in each merged region so sheet_to_json gets full data.
function expandMergedCells(sheet: XLSX.WorkSheet): void {
  const merges: XLSX.Range[] = sheet['!merges'] ?? [];
  for (const merge of merges) {
    const sourceAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const sourceCell = sheet[sourceAddr];
    if (!sourceCell) continue;
    for (let row = merge.s.r; row <= merge.e.r; row++) {
      for (let col = merge.s.c; col <= merge.e.c; col++) {
        if (row === merge.s.r && col === merge.s.c) continue;
        const addr = XLSX.utils.encode_cell({ r: row, c: col });
        if (!sheet[addr]) {
          sheet[addr] = { t: sourceCell.t, v: sourceCell.v, w: sourceCell.w };
        }
      }
    }
  }
}

// When a question spans multiple physical rows due to a multi-row merge (e.g.
// the Nithiodote question occupying Excel rows 57–58), the QA file ends up with
// more data rows than there are questions. This function returns only the FIRST
// row of each multi-row merge, discarding the continuation rows, so that the
// resulting array has exactly one row per question and S.No. → index is correct.
function readCollapsedQaRows(sheet: XLSX.WorkSheet): (string | null)[][] {
  const merges: XLSX.Range[] = sheet['!merges'] ?? [];

  // Collect data-row indices (0-based, after header) that are continuations
  // of a multi-row merge. Excel row r (0-based) → data index = r - 1.
  const skipDataRows = new Set<number>();
  for (const merge of merges) {
    if (merge.e.r > merge.s.r) {
      for (let excelRow = merge.s.r + 1; excelRow <= merge.e.r; excelRow++) {
        skipDataRows.add(excelRow - 1);
      }
    }
  }

  if (skipDataRows.size > 0) {
    console.log(
      `[Excel] Collapsing ${skipDataRows.size} continuation row(s) from multi-row merge(s):`,
      `data row(s) ${[...skipDataRows].sort((a, b) => a - b).map(i => i + 1).join(', ')}`
    );
  }

  // Fill merged cells so all rows in a merge carry the full values
  expandMergedCells(sheet);

  const allRows = XLSX.utils.sheet_to_json<(string | null)[]>(sheet, {
    header: 1, range: 1, defval: null,
  }) as (string | null)[][];

  // Drop continuation rows — keep only the first row of each multi-row merge
  return allRows.filter((_, idx) => !skipDataRows.has(idx));
}

// ---------------------------------------------------------------------------
// Excel readers
// ---------------------------------------------------------------------------

interface DoiEntry {
  doi: string;
  serialNo: number;  // S.No. from column A of the DOI file (1-based)
}

// Reads the DOI list file and returns {doi, serialNo} pairs.
// S.No. (col A) is used as the JOIN KEY to the QA file rows.
function readDoisFromExcel(filePath: string): DoiEntry[] {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  expandMergedCells(sheet);
  const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1, range: 1, defval: null,
  }) as (string | number | null)[][];

  const entries: DoiEntry[] = [];
  for (const row of rows) {
    const rawSNo = row[0];
    const doi    = row[1];
    if (typeof doi !== 'string' || !doi.startsWith('qna-')) continue;
    // S.No. column may be a cached number or a string from the =ROW()-1 formula
    const serialNo = typeof rawSNo === 'number'
      ? rawSNo
      : parseInt(String(rawSNo ?? '0'), 10);
    if (!isNaN(serialNo) && serialNo > 0) {
      entries.push({ doi, serialNo });
    }
  }

  console.log(`[Excel] Loaded ${entries.length} DOIs`);
  return entries;
}

// Maps doi → {serialNo, stem, answer} by joining S.No. (file 1) to
// the corresponding row in the QA file (S.No. 1 = row index 0, etc.).
// This is the authoritative join — it does NOT assume sequential DOI numbers.
function buildExpectedAnswerMap(
  entries: DoiEntry[],
  qaFile: string
): Map<string, QuestionData & { serialNo: number }> {
  const qaWorkbook = XLSX.readFile(qaFile);
  const qaSheet = qaWorkbook.Sheets[qaWorkbook.SheetNames[0]];
  // Use collapsed rows: multi-row merges are reduced to one row per question
  const qaRows = readCollapsedQaRows(qaSheet);

  const map = new Map<string, QuestionData & { serialNo: number }>();
  for (const { doi, serialNo } of entries) {
    const rowIndex = serialNo - 1;          // S.No. is 1-based → 0-based index
    const row = qaRows[rowIndex];
    if (!row) continue;
    // QA file columns: 0=Chapter, 1=Question No, 2=Question Stem, 3=Answers, 4=Solutions
    const stem   = String(row[2] ?? '').trim();
    const answer = String(row[3] ?? '').trim();
    if (stem || answer) map.set(doi, { serialNo, stem, answer });
  }

  console.log(`[Excel] Mapped ${map.size} expected answers`);
  return map;
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------

// Detects and attempts to pass a Cloudflare Turnstile challenge.
// Cloudflare embeds the widget in an iframe from challenges.cloudflare.com.
// In headed mode this usually succeeds; in headless mode it may not.
async function handleCloudflareChallenge(page: Page): Promise<void> {
  // Give the page a moment to inject the Turnstile iframe if it's going to
  await page.waitForTimeout(2000);

  const cfFrame = page.frames().find(f => f.url().includes('challenges.cloudflare.com'));

  if (!cfFrame) {
    // No Cloudflare iframe — challenge may be embedded directly on the page
    const directChallenge = await page.$('#content .cb-c, #content input[type="checkbox"]');
    if (!directChallenge) return; // No challenge detected
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

  // Wait for Cloudflare to emit a success signal or redirect away
  console.log('[Auth] Waiting for Turnstile verification...');
  try {
    await Promise.race([
      // Cloudflare redirects after success — wait for the URL to change
      page.waitForURL(url => !url.href.includes('challenges.cloudflare.com'), { timeout: 20000 }),
      // Or wait for the #success element to become visible on-page
      page.waitForFunction(
        () => {
          const s = document.querySelector('#success') as HTMLElement | null;
          return s !== null && s.style.display !== 'none' && s.style.visibility !== 'hidden';
        },
        { timeout: 20000 }
      ),
    ]);
    console.log('[Auth] Turnstile challenge passed.');
    // Give the redirect a moment to settle
    await page.waitForTimeout(1500);
  } catch {
    console.warn('[Auth] Turnstile did not resolve within 20 s.');
    console.warn('[Auth] TIP: Set HEADLESS=false and watch the browser, or switch to cookie auth (see .env).');
  }
}

async function applyAuth(context: BrowserContext): Promise<void> {
  // --- Cookie injection mode (most reliable — bypasses Cloudflare entirely) ---
  const authMethod = (process.env.AUTH_METHOD ?? 'login').trim().toLowerCase();

  if (authMethod === 'cookie') {
    type CookieParam = { name: string; value: string; domain: string; path: string; httpOnly: boolean; secure: boolean; sameSite: 'Strict' | 'Lax' | 'None' };
    const cookies: CookieParam[] = [];

    if (process.env.CF_CLEARANCE) {
      cookies.push({
        name: 'cf_clearance',
        value: process.env.CF_CLEARANCE,
        domain: '.pharmacylibrary.com',
        path: '/',
        httpOnly: false,
        secure: true,
        sameSite: 'None',
      });
    }
    if (process.env.SESSION_COOKIE_VALUE) {
      cookies.push({
        name: process.env.SESSION_COOKIE_NAME ?? 'JSESSIONID',
        value: process.env.SESSION_COOKIE_VALUE,
        domain: '.pharmacylibrary.com',
        path: '/',
        httpOnly: true,
        secure: true,
        sameSite: 'Lax',
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

  // --- Form login mode ---
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

    // Handle Cloudflare Turnstile before the login form is accessible
    await handleCloudflareChallenge(loginPage);

    // If Cloudflare redirected away from the login URL, navigate back to it
    if (!loginPage.url().includes('pharmacylibrary.com/action/login') &&
        !loginPage.url().includes('pharmacylibrary.com/login')) {
      console.log('[Auth] Navigating to login form after challenge...');
      await loginPage.goto('https://pharmacylibrary.com/action/login', {
        waitUntil: 'domcontentloaded',
        timeout: PAGE_TIMEOUT,
      });
    }

    // Wait for the login form to be present
    await loginPage.waitForSelector('input[name="login"], input[name="email"], input[type="email"]', {
      timeout: 10000,
    }).catch(() => {
      console.warn('[Auth] Login form not found — check AUTH_LOGIN_URL or inspect the page with HEADLESS=false');
    });

    // Fill credentials (Atypon uses name="login" for username)
    const usernameField = await loginPage.$('input[name="login"]') ??
                          await loginPage.$('input[name="email"]') ??
                          await loginPage.$('input[type="email"]');
    if (usernameField) await usernameField.fill(username);

    const passwordField = await loginPage.$('input[name="password"]') ??
                          await loginPage.$('input[type="password"]');
    if (passwordField) await passwordField.fill(password ?? '');

    await loginPage.click('button[type="submit"], input[type="submit"]');
    await loginPage.waitForNavigation({ waitUntil: 'networkidle', timeout: PAGE_TIMEOUT })
      .catch(() => { /* some sites don't navigate after login */ });

    const finalUrl = loginPage.url();
    if (finalUrl.includes('/login') || finalUrl.includes('/signin')) {
      console.warn('[Auth] WARNING: Still on login page — credentials may be wrong or Turnstile blocked submission');
      console.warn('[Auth] TIP: Try AUTH_METHOD=cookie in .env (see instructions at top of .env)');
    } else {
      console.log(`[Auth] Login succeeded → ${finalUrl}`);
    }
  } finally {
    await loginPage.close();
  }
}

// ---------------------------------------------------------------------------
// CSV helpers
// ---------------------------------------------------------------------------

const CSV_HEADER = [
  'doi', 'question_no', 'url', 'http_status',
  'has_question_text', 'page_stem',
  'reveal_button_present',
  'displayed_answer', 'expected_answer', 'answer_matches',
  'broken_equation_images', 'broken_other_images',
  'raw_mathml_visible', 'raw_entities_visible',
  'explanation_present', 'missing_alt_images',
  'console_errors', 'other_issues', 'notes',
].join(',');

function initCsvFile(csvPath: string): void {
  fs.mkdirSync(path.dirname(csvPath), { recursive: true });
  fs.writeFileSync(csvPath, CSV_HEADER + '\n', 'utf8');
}

function escapeCsv(value: string | boolean | number | null): string {
  if (value === null || value === undefined) return '';
  const s = String(value);
  if (s.includes(',') || s.includes('"') || s.includes('\n')) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function appendCsvRow(csvPath: string, r: PageResult): void {
  const row = [
    r.doi, r.question_no, r.url, r.http_status,
    r.has_question_text, r.page_stem,
    r.reveal_button_present,
    r.displayed_answer, r.expected_answer, r.answer_matches,
    r.broken_equation_images, r.broken_other_images,
    r.raw_mathml_visible, r.raw_entities_visible,
    r.explanation_present, r.missing_alt_images,
    r.console_errors, r.other_issues, r.notes,
  ].map(escapeCsv).join(',');
  fs.appendFileSync(csvPath, row + '\n', 'utf8');
}

// ---------------------------------------------------------------------------
// Answer comparison — handles format variations between Excel and site
// ---------------------------------------------------------------------------

function normalise(s: string): string {
  return s.replace(/\s+/g, ' ').trim().toLowerCase();
}

// Extracts a numeric value from a string, stripping units/symbols.
function extractNumber(s: string): number | null {
  const m = s.match(/[\d]+(?:[.,]\d+)?/);
  if (!m) return null;
  return parseFloat(m[0].replace(',', '.'));
}

function answersMatch(displayed: string, expected: string): boolean {
  const d = normalise(displayed);
  const e = normalise(expected);
  if (d === e) return true;

  // Percentage vs decimal: "4.2%" ↔ "0.042"
  const dPct = d.match(/^([\d.]+)%$/);
  const ePct = e.match(/^([\d.]+)%$/);
  const dNum = extractNumber(d);
  const eNum = extractNumber(e);

  if (dPct && eNum !== null) {
    if (Math.abs(parseFloat(dPct[1]) / 100 - eNum) < 0.00001) return true;
  }
  if (ePct && dNum !== null) {
    if (Math.abs(parseFloat(ePct[1]) / 100 - dNum) < 0.00001) return true;
  }

  // Same number, different units or formatting (e.g. "1.4 g" vs "1.4g")
  if (dNum !== null && eNum !== null && Math.abs(dNum - eNum) < 0.00001) {
    // Only pass if the unit words also roughly match
    const dWords = d.replace(/[\d.,]/g, '').trim();
    const eWords = e.replace(/[\d.,]/g, '').trim();
    if (dWords === eWords || dWords.includes(eWords) || eWords.includes(dWords)) return true;
  }

  return false;
}

// ---------------------------------------------------------------------------
// Per-page checks
// ---------------------------------------------------------------------------

async function checkQuestionContent(
  page: Page,
  result: PageResult,
  expected: QuestionData | undefined
): Promise<void> {
  // 1. Question stem
  const stemEl = await page.$(SELECTORS.questionStem);
  if (!stemEl) {
    result.has_question_text = false;
    result.page_stem = '';
    result.hasIssues = true;
    result.notes += ' | question stem (#question p) not found';
  } else {
    const stemText = (await stemEl.innerText()).trim();
    result.page_stem = stemText;
    result.has_question_text = stemText.length > 0;
    if (!result.has_question_text) {
      result.hasIssues = true;
      result.notes += ' | question stem is empty';
    }
  }

  // 2. Reveal button
  const revealBtn = await page.$(SELECTORS.revealButton);
  result.reveal_button_present = revealBtn !== null;
  if (!revealBtn) {
    result.hasIssues = true;
    result.notes += ' | reveal radio button (input[name="choice"]) not found';
    return;
  }

  // 3. Click to reveal, then wait for answer block AND any images inside it to load
  await revealBtn.click().catch(() => {});

  await page.waitForFunction(
    () => {
      const el = document.querySelector('#answer-block');
      return el !== null && !el.classList.contains('hide');
    },
    { timeout: 8000 }
  ).catch(() => {});

  // Give images inside the answer block time to finish loading
  await page.waitForFunction(
    () => {
      const imgs = Array.from(document.querySelectorAll('#answer-block img'));
      return imgs.length === 0 || imgs.every(img => (img as HTMLImageElement).complete);
    },
    { timeout: 10000 }
  ).catch(() => {});

  // 4. Read displayed answer
  const answerEl = await page.$(SELECTORS.answerText);
  result.displayed_answer = answerEl
    ? (await answerEl.innerText()).trim()
    : '';

  if (!result.displayed_answer) {
    result.hasIssues = true;
    result.notes += ' | answer text not found after reveal';
  }

  // 5. Compare with expected answer from Excel
  if (expected?.answer) {
    result.expected_answer = expected.answer;
    result.answer_matches = answersMatch(result.displayed_answer, expected.answer);
    if (!result.answer_matches) {
      result.hasIssues = true;
      result.notes += ` | answer mismatch: got "${result.displayed_answer}" expected "${expected.answer}"`;
    }
  } else {
    result.answer_matches = null;
  }

  // 6. Explanation present — accept text OR images (equation-only explanations are valid)
  const explanationEl = await page.$(SELECTORS.explanation);
  if (explanationEl) {
    const hasText = (await explanationEl.innerText()).trim().length > 0;
    const hasImages = await explanationEl.$('img') !== null;
    result.explanation_present = hasText || hasImages;
  } else {
    result.explanation_present = false;
  }
  if (!result.explanation_present) {
    result.hasIssues = true;
    result.notes += ' | explanation missing';
  }
}

async function checkImagesAndText(page: Page, result: PageResult): Promise<void> {
  // Equation images (.qna-equation img) — these are broken if naturalWidth === 0
  const imageData = await page.evaluate((): { eqBroken: string[]; otherBroken: string[]; noAlt: string[] } => {
    const eqBroken: string[] = [];
    const otherBroken: string[] = [];
    const noAlt: string[] = [];

    document.querySelectorAll('img').forEach(img => {
      const isBroken = img.complete && img.naturalWidth === 0 && !!img.src;
      const isEq = img.closest('.qna-equation') !== null;
      if (isBroken) (isEq ? eqBroken : otherBroken).push(img.src);
      if (!img.hasAttribute('alt')) noAlt.push(img.src || '[no-src]');
    });

    return { eqBroken, otherBroken, noAlt };
  });

  result.broken_equation_images = imageData.eqBroken.join(' | ');
  result.broken_other_images    = imageData.otherBroken.join(' | ');
  result.missing_alt_images     = imageData.noAlt.join(' | ');

  if (imageData.eqBroken.length > 0) {
    result.hasIssues = true;
    result.notes += ` | ${imageData.eqBroken.length} broken equation image(s)`;
  }
  if (imageData.otherBroken.length > 0) {
    result.hasIssues = true;
    result.notes += ` | ${imageData.otherBroken.length} broken other image(s)`;
  }

  // Raw MathML or HTML entities visible in rendered text
  const bodyText = await page.evaluate(() => document.body.innerText);
  result.raw_mathml_visible   = /<mml:|<\/mml:/i.test(bodyText);
  result.raw_entities_visible = /&amp;|&lt;|&gt;|&#x[0-9a-fA-F]+;|&#\d+;/.test(bodyText);

  if (result.raw_mathml_visible) {
    result.hasIssues = true;
    result.notes += ' | raw MathML tags visible in text';
  }
  if (result.raw_entities_visible) {
    result.hasIssues = true;
    result.notes += ' | raw HTML entities visible in text';
  }

  // Unicode replacement characters
  if (/�/.test(bodyText)) {
    const count = (bodyText.match(/�/g) ?? []).length;
    result.hasIssues = true;
    result.other_issues += (result.other_issues ? ' | ' : '') + `replacement_chars:${count}`;
    result.notes += ` | ${count} Unicode replacement char(s)`;
  }
}

// ---------------------------------------------------------------------------
// Cloudflare challenge detector
// ---------------------------------------------------------------------------

// Returns true if the current page is a Cloudflare security check.
// Checks the page title first (fast), then falls back to body text scan.
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

// Waits up to 30 s for a Cloudflare challenge to auto-resolve,
// then reloads the original URL. Returns the new HTTP status.
async function waitForChallengeAndReload(page: Page, url: string): Promise<number> {
  console.warn(`\n[CF] Challenge detected — waiting up to 30 s for auto-resolve...`);
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
  ).catch(() => {
    console.warn('[CF] Challenge did not resolve in 30 s — retrying navigation.');
  });

  // Small buffer after the challenge resolves
  await page.waitForTimeout(1500);

  const response = await page.goto(url, {
    waitUntil: 'domcontentloaded',
    timeout: PAGE_TIMEOUT,
  }).catch(() => null);

  return response?.status() ?? 0;
}

// ---------------------------------------------------------------------------
// Main page checker
// ---------------------------------------------------------------------------

async function checkPage(
  page: Page,
  doi: string,
  expected: QuestionData | undefined
): Promise<PageResult> {
  const url = `${BASE_URL}/do/${DOI_PREFIX}/${doi}/abs/`;
  const result: PageResult = {
    doi,
    question_no: expected?.serialNo ?? null,
    url,
    http_status: 0,
    has_question_text: null,
    page_stem: '',
    reveal_button_present: null,
    displayed_answer: '',
    expected_answer: expected?.answer ?? '',
    answer_matches: null,
    broken_equation_images: '',
    broken_other_images: '',
    raw_mathml_visible: false,
    raw_entities_visible: false,
    explanation_present: null,
    missing_alt_images: '',
    console_errors: '',
    other_issues: '',
    notes: '',
    hasIssues: false,
  };

  const consoleErrors: string[] = [];
  const onConsole = (msg: ConsoleMessage) => {
    if (msg.type() === 'error') consoleErrors.push(msg.text());
  };
  page.on('console', onConsole);

  try {
    const response = await page.goto(url, {
      waitUntil: 'domcontentloaded',
      timeout: PAGE_TIMEOUT,
    });

    result.http_status = response?.status() ?? 0;

    // --- Cloudflare challenge check (run immediately after load) ---
    if (await isCloudflareChallenge(page)) {
      result.http_status = await waitForChallengeAndReload(page, url);
      // If still a challenge after waiting, flag and skip
      if (await isCloudflareChallenge(page)) {
        result.hasIssues = true;
        result.notes = 'Cloudflare challenge not resolved — refresh login session and rerun';
        const screenshotPath = path.join(SCREENSHOTS_DIR, `${doi}.png`);
        await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
        result.notes += ` | screenshot: ${doi}.png`;
        return result;
      }
      console.warn(`[CF] Challenge resolved — continuing checks for ${doi}`);
    }

    if (result.http_status === 404) {
      result.notes = 'Page not found (404)';
      return result;
    }
    if (result.http_status !== 200) {
      result.hasIssues = true;
      result.notes += `non-200 HTTP status: ${result.http_status}`;
    }

    await checkQuestionContent(page, result, expected);
    await checkImagesAndText(page, result);

    // Console errors (filter out noisy Cloudflare/infra errors)
    const realErrors = consoleErrors.filter(e =>
      !e.includes('challenges.cloudflare') &&
      !e.includes('favicon')
    );
    result.console_errors = realErrors.join(' | ');

    if (result.hasIssues) {
      const screenshotPath = path.join(SCREENSHOTS_DIR, `${doi}.png`);
      await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
      result.notes += ` | screenshot: ${doi}.png`;
    }

  } catch (err: unknown) {
    result.hasIssues = true;
    result.notes = `Error: ${(err as Error).message}`;
  } finally {
    page.off('console', onConsole);
  }

  return result;
}

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

function printSummary(stats: SummaryStats): void {
  console.log('\n' + '='.repeat(60));
  console.log('QA SUMMARY');
  console.log('='.repeat(60));
  console.log(`Total pages checked        : ${stats.total}`);
  console.log(`Pages with issues          : ${stats.issues}`);
  console.log(`404 not found              : ${stats.notFound}`);
  console.log('-'.repeat(60));
  console.log('Issue breakdown:');
  console.log(`  Missing question text    : ${stats.missingQuestionText}`);
  console.log(`  Missing reveal button    : ${stats.missingRevealButton}`);
  console.log(`  Answer mismatches        : ${stats.answerMismatches}`);
  console.log(`  Broken equation images   : ${stats.brokenEquationImages}`);
  console.log(`  Broken other images      : ${stats.brokenOtherImages}`);
  console.log(`  Raw MathML visible       : ${stats.rawMathml}`);
  console.log(`  Raw HTML entities        : ${stats.rawEntities}`);
  console.log(`  Missing explanation      : ${stats.missingExplanation}`);
  console.log(`  JS console errors        : ${stats.consoleErrors}`);
  console.log(`  Unicode replacement chars: ${stats.unicodeReplacement}`);
  console.log('='.repeat(60));
  console.log('Report      : output/qna_qa_report.csv');
  console.log('Screenshots : screenshots/');
}

function updateStats(stats: SummaryStats, result: PageResult): void {
  stats.total++;
  if (result.http_status === 404) stats.notFound++;
  if (result.hasIssues) stats.issues++;
  if (result.has_question_text === false) stats.missingQuestionText++;
  if (result.reveal_button_present === false) stats.missingRevealButton++;
  if (result.answer_matches === false) stats.answerMismatches++;
  if (result.broken_equation_images) stats.brokenEquationImages++;
  if (result.broken_other_images) stats.brokenOtherImages++;
  if (result.raw_mathml_visible) stats.rawMathml++;
  if (result.raw_entities_visible) stats.rawEntities++;
  if (result.explanation_present === false) stats.missingExplanation++;
  if (result.console_errors) stats.consoleErrors++;
  if (result.other_issues.includes('replacement_chars')) stats.unicodeReplacement++;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  console.log('NAPLEX QA Script — staging.pharmacylibrary.com');
  console.log('='.repeat(60));

  const doiEntries = readDoisFromExcel(EXCEL_DOI_PATH);
  const expectedAnswers = buildExpectedAnswerMap(doiEntries, EXCEL_QA_PATH);
  const dois = doiEntries.map(e => e.doi);
  if (dois.length === 0) {
    console.error('[Error] No DOIs found. Check Excel file paths.');
    process.exit(1);
  }

  initCsvFile(CSV_PATH);
  fs.mkdirSync(SCREENSHOTS_DIR, { recursive: true });

  const isHeadless   = process.env.HEADLESS      !== 'false';
  const debugPause   = process.env.DEBUG_PAUSE   === 'true';
  const browserMode  = (process.env.BROWSER_MODE ?? 'launch').trim().toLowerCase();
  console.log(`[Browser] mode=${browserMode}  headless=${isHeadless}  debugPause=${debugPause}`);

  let browser: Browser | null = null;
  let cdpBrowser: Browser | null = null; // CDP connection — we don't own/close this
  let context: BrowserContext;
  let browserAlive = true;

  if (browserMode === 'cdp') {
    const cdpUrl = process.env.CDP_URL ?? 'http://localhost:9222';
    console.log(`[Browser] Connecting to your running Edge via CDP: ${cdpUrl}`);
    console.log('[Browser] Make sure open-edge.bat is running and you are logged in.');
    cdpBrowser = await chromium.connectOverCDP(cdpUrl);
    const existingContexts = cdpBrowser.contexts();
    if (existingContexts.length === 0) {
      throw new Error('No browser contexts found — is Edge open and logged in? Run open-edge.bat first.');
    }
    context = existingContexts[0];
    // Listen for disconnect (user closes Edge)
    cdpBrowser.on('disconnected', () => { browserAlive = false; });

  } else {
    browser = await chromium.launch({ headless: isHeadless });
    context = await browser.newContext({ ignoreHTTPSErrors: true });
    await applyAuth(context);
  }

  // In CDP mode we open a new tab inside the user's browser — close just that tab at the end
  const page = await context.newPage();

  // Navigate to the first QA page so the Inspector opens with useful context
  if (debugPause) {
    const firstUrl = `${BASE_URL}/do/${DOI_PREFIX}/${dois[0]}/abs/`;
    console.log(`[Debug] Navigating to first page before pause: ${firstUrl}`);
    await page.goto(firstUrl, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT })
      .catch(() => {}); // ignore nav errors — user can inspect the 403 itself
    console.log('[Debug] Pausing — use Playwright Inspector to explore the page.');
    console.log('        Press Resume (▶) in the Inspector to start the sweep.');
    await page.pause();
  }

  const stats: SummaryStats = {
    total: 0, issues: 0, notFound: 0,
    missingQuestionText: 0, missingRevealButton: 0, answerMismatches: 0,
    brokenEquationImages: 0, brokenOtherImages: 0,
    rawMathml: 0, rawEntities: 0, missingExplanation: 0,
    consoleErrors: 0, unicodeReplacement: 0,
  };

  if (browser) browser.on('disconnected', () => { browserAlive = false; });
  context.on('close', () => { browserAlive = false; });

  console.log(`\nStarting QA sweep — ${dois.length} DOIs\n`);

  for (let i = 0; i < dois.length; i++) {
    if (!browserAlive) {
      console.warn('\n[Stopped] Browser was closed — saving partial results.');
      break;
    }

    const doi = dois[i];
    const progress = `[${String(i + 1).padStart(3)}/${dois.length}]`;

    let result: PageResult;
    try {
      result = await checkPage(page, doi, expectedAnswers.get(doi));
    } catch (err: unknown) {
      const msg = (err as Error).message ?? String(err);
      // Browser or page was closed mid-check — save what we have and exit loop
      if (msg.includes('closed') || msg.includes('disconnected') || msg.includes('Target')) {
        console.warn(`\n[Stopped] Browser closed during ${doi} — saving partial results.`);
        browserAlive = false;
        break;
      }
      result = {
        doi, question_no: null, url: `${BASE_URL}/do/${DOI_PREFIX}/${doi}/abs/`,
        http_status: 0, has_question_text: null, page_stem: '',
        reveal_button_present: null, displayed_answer: '', expected_answer: '',
        answer_matches: null, broken_equation_images: '', broken_other_images: '',
        raw_mathml_visible: false, raw_entities_visible: false,
        explanation_present: null, missing_alt_images: '',
        console_errors: '', other_issues: '', hasIssues: true,
        notes: `Unhandled error: ${msg}`,
      };
    }

    updateStats(stats, result);
    appendCsvRow(CSV_PATH, result);

    const statusLine = result.http_status === 404
      ? '404 NOT FOUND'
      : result.hasIssues
        ? `ISSUES: ${result.notes.trim()}`
        : 'OK';
    console.log(`${progress} ${doi}  ${statusLine}`);

    if (browserAlive && i < dois.length - 1) {
      await page.waitForTimeout(DELAY_BETWEEN_PAGES).catch(() => {});
    }
  }

  if (browserAlive) {
    if (cdpBrowser) {
      // CDP mode: close just our tab, leave the user's browser running
      await page.close().catch(() => {});
    } else {
      await context.close().catch(() => {});
      if (browser) await browser.close().catch(() => {});
    }
  }
  printSummary(stats);
}

main().catch(err => {
  console.error('[Fatal]', err);
  process.exit(1);
});
