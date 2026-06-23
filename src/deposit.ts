/**
 * deposit.ts — Atypon (pharmacylibrary.com WAT console) → Impelsys deposit automation.
 *
 * Independent of qa-check.ts / find-dois.ts. Uses a single browser, reusing the same
 * Edge debug profile created by open-edge.bat (CDP mode), so the manual login is shared.
 *
 * Flow (per run):
 *   1. Connect to the running Edge (CDP) and open a page.
 *   2. Navigate to https://pharmacylibrary.com/wat (Atypon WAT console).
 *   3. Click the side-panel button → open the console.
 *   4. Click the "Backstage" tab.
 *   5. Select the environment ("live") in the first combo-box.
 *   6. Select the publication type (Digital Object Publication [default] / Books) in the combo-box.
 *   7. For each not-yet-migrated DOI from the DOs tab of Migration_record.xlsx (filtered by Category):
 *        a. Type the DOI into the "pubDOI" field and search.
 *        b. Find the matching row in the grid, right-click it, click "Deposit".
 *        c. In the deposit dialog, tick "Impelsys Migration Feed".
 *        d. Wait for the "Deposit" button to enable, click it.
 *        e. Confirm "Yes".
 *        f. Wait for "Please check submissions page later for result." (deposit started).
 *        g. Mark the row Status = "started", Date of attempt = today, in the Excel file.
 *        h. Close the confirmation, wait DEPOSIT_DELAY_MS, continue.
 *
 * Resume: rows whose Status column is already non-blank are skipped, so re-running
 * continues where it left off. A JSON checkpoint (output/deposit-progress.json) mirrors
 * progress in case the Excel file is locked (open in Excel) during a write.
 *
 * Single-DOI mode: set DEPOSIT_SINGLE_DOI to deposit exactly one DOI.
 */

import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { chromium, Browser, BrowserContext, Page, Locator } from 'playwright';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const ROOT = path.resolve(__dirname, '..');

function envStr(name: string, def = ''): string {
  const v = process.env[name];
  return v === undefined || v === null ? def : v.trim();
}
function envInt(name: string, def: number): number {
  const v = parseInt(envStr(name), 10);
  return Number.isFinite(v) ? v : def;
}
function envBool(name: string, def = false): boolean {
  const v = envStr(name).toLowerCase();
  if (v === '') return def;
  return v === 'true' || v === '1' || v === 'yes';
}

const CONFIG = {
  // --- Excel source of truth ---
  xlsxPath: path.resolve(ROOT, envStr('DEPOSIT_RECORD_XLSX', 'Data/DOIs record/Migration_record.xlsx')),
  sheet: envStr('DEPOSIT_SHEET', 'DOs'),
  // Which Category column values to process: NAPLEX | other | all
  category: envStr('DEPOSIT_CATEGORY', 'NAPLEX').toLowerCase(),
  // App publication-type dropdown: "digital" → Digital Object Publication, "books" → Books
  publicationType: envStr('DEPOSIT_PUBLICATION_TYPE', 'digital').toLowerCase(),
  // App environment dropdown option (first combo-box)
  envOption: envStr('DEPOSIT_ENV_OPTION', 'live'),
  // Optional: deposit exactly one DOI and stop (overrides the Excel-driven loop selection)
  singleDoi: envStr('DEPOSIT_SINGLE_DOI', ''),
  // Optional cap on number of deposits this run (0 = no limit)
  limit: envInt('DEPOSIT_LIMIT', 0),
  // Delay between two deposits (ms)
  delayMs: envInt('DEPOSIT_DELAY_MS', 5000),
  // Value written to the Status column on a started deposit
  statusValue: envStr('DEPOSIT_STATUS_VALUE', 'started'),
  // If true, also write "error: <reason>" to the Status column on failure (so it is skipped next run)
  markErrors: envBool('DEPOSIT_MARK_ERRORS', false),
  // If true, do everything EXCEPT the final "Yes" confirmation (safe trial of selectors)
  dryRun: envBool('DEPOSIT_DRY_RUN', false),
  // If true, open the Playwright Inspector right after /wat loads so you can pick locators.
  debugPause: envBool('DEBUG_PAUSE', false),

  // --- Browser ---
  browserMode: envStr('BROWSER_MODE', 'cdp').toLowerCase(),
  cdpPort: envInt('CDP_BASE_PORT', 9222),
  headless: !(envStr('HEADLESS', 'false') === 'false'),
  userDataDir: path.resolve(ROOT, envStr('DEPOSIT_USER_DATA_DIR', '.edge-debug-profile-0')),

  // --- Timeouts (ms) ---
  navTimeout: envInt('DEPOSIT_NAV_TIMEOUT', 60_000),
  selTimeout: envInt('DEPOSIT_SEL_TIMEOUT', 30_000),
  gridTimeout: envInt('DEPOSIT_GRID_TIMEOUT', 20_000),
  depositEnableTimeout: envInt('DEPOSIT_ENABLE_TIMEOUT', 30_000),

  // How long to wait for the optional OK/Yes confirmation after clicking Deposit (ms).
  confirmTimeout: envInt('DEPOSIT_CONFIRM_TIMEOUT', 6000),
};

// ---------------------------------------------------------------------------
// Hard-coded UI controls (Atypon WAT console). These are intentionally NOT in .env —
// they identify fixed elements of the page. Edit here if the UI itself changes.
// ---------------------------------------------------------------------------

const UI = {
  sidePanelButton: 'Administration', // side-panel button that opens the console
  envComboText: 'Area', // text identifying the environment combo-box
  typeComboText: 'Type', // text identifying the publication-type combo-box
  publicationLabels: {
    digital: 'Digital Object Publications',
    books: 'Books',
  } as Record<string, string>,
};

// A normal run must never open the Playwright Inspector. Only DEBUG_PAUSE=true enables it;
// otherwise drop any PWDEBUG left over in the shell so it can't force the Inspector open.
if (!CONFIG.debugPause) {
  delete process.env.PWDEBUG;
  delete process.env.PLAYWRIGHT_PAUSE;
}

const BASE_URL = 'https://pharmacylibrary.com';
const WAT_URL = `${BASE_URL}/wat`;

const SHOT_DIR = path.join(ROOT, 'output', 'deposit-screenshots');
const PROGRESS_FILE = path.join(ROOT, 'output', 'deposit-progress.json');

const PUBLICATION_LABEL =
  CONFIG.publicationType === 'books' ? UI.publicationLabels.books : UI.publicationLabels.digital;

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------

function ts(): string {
  return new Date().toISOString().replace('T', ' ').slice(0, 19);
}
function log(msg: string): void {
  console.log(`[${ts()}] ${msg}`);
}
function warn(msg: string): void {
  console.warn(`[${ts()}] [WARN] ${msg}`);
}

/** Human-friendly date matching the existing sheet style, e.g. "2026 June 23". */
function todayLabel(): string {
  const d = new Date();
  const month = d.toLocaleString('en-US', { month: 'long' });
  return `${d.getFullYear()} ${month} ${d.getDate()}`;
}

// ---------------------------------------------------------------------------
// Excel record store
// ---------------------------------------------------------------------------

interface DoiRow {
  rowIdx: number; // 0-based sheet row index (header is 0, first data row is 1)
  doi: string;
  status: string;
  date: string;
  category: string;
}

const COL = { DOI: 0, STATUS: 1, DATE: 2, CATEGORY: 3 } as const;

class RecordStore {
  private wb: XLSX.WorkBook;
  private ws: XLSX.WorkSheet;
  rows: DoiRow[] = [];

  constructor(private filePath: string, private sheetName: string) {
    if (!fs.existsSync(filePath)) {
      throw new Error(`Record file not found: ${filePath}`);
    }
    this.wb = XLSX.readFile(filePath, { cellStyles: true });
    const ws = this.wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet "${sheetName}" not found in ${filePath}`);
    this.ws = ws;

    const grid = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, blankrows: false, defval: '' });
    for (let r = 1; r < grid.length; r++) {
      const row = grid[r];
      const doi = String(row[COL.DOI] ?? '').trim();
      if (!doi) continue;
      this.rows.push({
        rowIdx: r,
        doi,
        status: String(row[COL.STATUS] ?? '').trim(),
        date: String(row[COL.DATE] ?? '').trim(),
        category: String(row[COL.CATEGORY] ?? '').trim(),
      });
    }
  }

  findByDoi(doi: string): DoiRow | undefined {
    const needle = doi.trim().toLowerCase();
    return this.rows.find((r) => r.doi.toLowerCase() === needle);
  }

  /** Update Status + Date cells for a row and persist to disk. Returns false if the write failed. */
  setStatus(row: DoiRow, status: string, date: string): boolean {
    row.status = status;
    row.date = date;
    this.setCell(row.rowIdx, COL.STATUS, status);
    this.setCell(row.rowIdx, COL.DATE, date);
    return this.save();
  }

  private setCell(r: number, c: number, value: string): void {
    const addr = XLSX.utils.encode_cell({ r, c });
    this.ws[addr] = { t: 's', v: value };
    // Make sure the sheet range covers the cell (columns already exist, so this is a safety net).
    const ref = this.ws['!ref'];
    if (ref) {
      const range = XLSX.utils.decode_range(ref);
      if (c > range.e.c) range.e.c = c;
      if (r > range.e.r) range.e.r = r;
      this.ws['!ref'] = XLSX.utils.encode_range(range);
    }
  }

  save(): boolean {
    try {
      XLSX.writeFile(this.wb, this.filePath);
      return true;
    } catch (err) {
      warn(`Could not write Excel (${(err as Error).message}). Is the file open in Excel? Progress is still in ${path.basename(PROGRESS_FILE)}.`);
      return false;
    }
  }
}

// ---------------------------------------------------------------------------
// JSON progress checkpoint (backup for resume when Excel is locked)
// ---------------------------------------------------------------------------

interface Progress {
  completed: Record<string, { status: string; date: string; at: string }>;
  failed: Record<string, { reason: string; at: string }>;
}

function loadProgress(): Progress {
  if (!fs.existsSync(PROGRESS_FILE)) return { completed: {}, failed: {} };
  try {
    return JSON.parse(fs.readFileSync(PROGRESS_FILE, 'utf-8')) as Progress;
  } catch {
    return { completed: {}, failed: {} };
  }
}

function saveProgress(p: Progress): void {
  fs.mkdirSync(path.dirname(PROGRESS_FILE), { recursive: true });
  fs.writeFileSync(PROGRESS_FILE, JSON.stringify(p, null, 2), 'utf-8');
}

// ---------------------------------------------------------------------------
// Browser setup
// ---------------------------------------------------------------------------

interface BrowserHandle {
  browser: Browser | null; // null in CDP mode
  context: BrowserContext;
  page: Page;
  close: () => Promise<void>;
}

async function openBrowser(): Promise<BrowserHandle> {
  if (CONFIG.browserMode === 'cdp') {
    const cdpUrl = `http://localhost:${CONFIG.cdpPort}`;
    log(`Connecting to Edge via CDP at ${cdpUrl} (run open-edge.bat first and log in).`);
    const browser = await chromium.connectOverCDP(cdpUrl);
    const contexts = browser.contexts();
    if (contexts.length === 0) {
      throw new Error(`No browser context on port ${CONFIG.cdpPort}. Run: open-edge.bat 1`);
    }
    const context = contexts[0];
    const page = context.pages()[0] ?? (await context.newPage());
    return {
      browser: null,
      context,
      page,
      close: async () => {
        // Do NOT close the user's Edge window in CDP mode; just disconnect.
        await browser.close().catch(() => {});
      },
    };
  }

  // launch (persistent) mode — reuse the same profile dir so login persists.
  log(`Launching persistent Edge profile at ${CONFIG.userDataDir} (headless=${CONFIG.headless}).`);
  const context = await chromium.launchPersistentContext(CONFIG.userDataDir, {
    headless: CONFIG.headless,
    channel: 'msedge',
    ignoreHTTPSErrors: true,
    viewport: null,
  });
  const page = context.pages()[0] ?? (await context.newPage());
  return {
    browser: null,
    context,
    page,
    close: async () => {
      await context.close().catch(() => {});
    },
  };
}

// ---------------------------------------------------------------------------
// Interaction helpers (Playwright pierces open shadow DOM with CSS / text engines)
// ---------------------------------------------------------------------------

async function settle(page: Page, ms = 800): Promise<void> {
  await page.waitForLoadState('networkidle', { timeout: 15_000 }).catch(() => {});
  await page.waitForTimeout(ms);
}

async function clickWhenReady(loc: Locator, label: string, timeout = CONFIG.selTimeout): Promise<void> {
  log(`  · click: ${label}`);
  await loc.first().waitFor({ state: 'visible', timeout });
  await loc.first().scrollIntoViewIfNeeded().catch(() => {});
  await loc.first().click({ timeout });
}

/**
 * Open the combo-box identified by `comboText` (part of its label/text) and pick the option
 * whose accessible name matches `optionText` (case-insensitive substring).
 */
async function selectComboOption(page: Page, comboText: string, optionText: string): Promise<void> {
  log(`  · combo [${comboText}] → "${optionText}"`);
  const combo = page.locator('vaadin-combo-box').filter({ hasText: comboText }).first();
  await combo.waitFor({ state: 'visible', timeout: CONFIG.selTimeout });
  await combo.locator('#toggleButton').click();
  const option = page.getByRole('option', { name: optionText }).first();
  await option.waitFor({ state: 'visible', timeout: CONFIG.selTimeout });
  await option.click();
  await page.waitForTimeout(500);
}

/** Step 3 + 4: click the side-panel ("Administration") button, then the Backstage tab. */
async function openBackstage(page: Page): Promise<void> {
  const sideBtn = page.getByRole('button', { name: UI.sidePanelButton });
  await clickWhenReady(sideBtn, `side-panel button (${UI.sidePanelButton})`);
  await settle(page);

  const backstage = page.locator('span').filter({ hasText: 'Backstage' }).first();
  await clickWhenReady(backstage, 'Backstage tab');
  await settle(page);
}

/** Type the DOI into the "DOI" (pubDOI) search field and trigger the search. */
async function searchDoi(page: Page, doi: string): Promise<void> {
  let input = page.getByRole('textbox', { name: 'DOI' }).first();
  if (!(await input.isVisible().catch(() => false))) {
    input = page.locator('input[name="pubDOI"]').first();
  }
  await input.waitFor({ state: 'visible', timeout: CONFIG.selTimeout });
  await input.click();
  await input.fill('');
  await input.fill(doi);
  await input.press('Enter');
  await settle(page, 1200);
}

/** Find the grid row whose accessible name contains `doi`, right-click it, click "Deposit". */
async function openDepositMenu(page: Page, doi: string): Promise<void> {
  const cell = page.getByRole('gridcell', { name: doi }).first();
  await cell.waitFor({ state: 'visible', timeout: CONFIG.gridTimeout });
  await cell.scrollIntoViewIfNeeded().catch(() => {});
  await cell.click({ button: 'right' });

  const deposit = page.getByRole('menuitem', { name: 'Deposit' }).first();
  await deposit.waitFor({ state: 'visible', timeout: CONFIG.selTimeout });
  await deposit.click();
  await settle(page);
}

/** In the deposit dialog: tick "Impelsys Migration Feed", click Deposit, confirm OK/Yes if asked. */
async function confirmDeposit(page: Page): Promise<void> {
  const checkbox = page.getByRole('checkbox', { name: 'Impelsys Migration Feed' }).first();
  await checkbox.waitFor({ state: 'visible', timeout: CONFIG.selTimeout });
  await checkbox.check().catch(async () => {
    await checkbox.click().catch(() => {});
  });

  // The Deposit button auto-enables once the feed is ticked; Playwright waits for "enabled".
  const depositBtn = page.getByRole('button', { name: 'Deposit', exact: true }).first();
  await depositBtn.waitFor({ state: 'visible', timeout: CONFIG.selTimeout });

  if (CONFIG.dryRun) {
    log('  · DRY RUN — reached the Deposit button; not depositing.');
    await page.keyboard.press('Escape').catch(() => {});
    throw new DryRunStop();
  }

  await clickWhenReady(depositBtn, 'Deposit (dialog)', CONFIG.depositEnableTimeout);

  // Optional confirmation dialog — sometimes an "OK"/"Yes" button appears, sometimes not.
  const confirm = page.getByRole('button', { name: /^(ok|yes)$/i }).first();
  try {
    await confirm.waitFor({ state: 'visible', timeout: CONFIG.confirmTimeout });
    await clickWhenReady(confirm, 'confirm (OK/Yes)');
  } catch {
    log('  · no confirmation dialog (continuing)');
  }
}

class DryRunStop extends Error {
  constructor() {
    super('dry-run');
    this.name = 'DryRunStop';
  }
}

/** Wait for the "deposit started" confirmation notification, then close it. */
async function awaitStartedAndClose(page: Page): Promise<void> {
  const started = page
    .getByText('Please check submissions page later for result', { exact: false })
    .first();
  const note = page.locator('.dialog-notification').first();
  await Promise.race([
    started.waitFor({ state: 'visible', timeout: CONFIG.selTimeout }),
    note.waitFor({ state: 'visible', timeout: CONFIG.selTimeout }),
  ]);
  log('  · deposit started ✓');

  const close = page.getByRole('button', { name: 'Close' }).first();
  await close.waitFor({ state: 'visible', timeout: CONFIG.selTimeout }).catch(() => {});
  await close.click().catch(() => {});
  await settle(page, 600);
}

// ---------------------------------------------------------------------------
// Per-DOI deposit
// ---------------------------------------------------------------------------

async function depositOne(page: Page, doi: string): Promise<void> {
  await searchDoi(page, doi);
  await openDepositMenu(page, doi);
  await confirmDeposit(page);
  await awaitStartedAndClose(page);
}

async function screenshot(page: Page, name: string): Promise<void> {
  try {
    fs.mkdirSync(SHOT_DIR, { recursive: true });
    const safe = name.replace(/[^a-z0-9._-]/gi, '_');
    await page.screenshot({ path: path.join(SHOT_DIR, `${safe}.png`), fullPage: false });
  } catch {
    /* best-effort */
  }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  console.log('='.repeat(64));
  console.log(' Atypon → Impelsys deposit automation');
  console.log('='.repeat(64));
  log(`Record:        ${CONFIG.xlsxPath} [${CONFIG.sheet}]`);
  log(`Category:      ${CONFIG.category}`);
  log(`Publication:   ${PUBLICATION_LABEL}`);
  log(`Environment:   ${CONFIG.envOption}`);
  log(`Delay:         ${CONFIG.delayMs} ms`);
  log(`Dry run:       ${CONFIG.dryRun}`);
  if (CONFIG.singleDoi) log(`Single DOI:    ${CONFIG.singleDoi}`);

  const store = new RecordStore(CONFIG.xlsxPath, CONFIG.sheet);
  const progress = loadProgress();

  // Build the work list.
  let work: DoiRow[];
  if (CONFIG.singleDoi) {
    const existing = store.findByDoi(CONFIG.singleDoi);
    work = [existing ?? { rowIdx: -1, doi: CONFIG.singleDoi, status: '', date: '', category: '' }];
  } else {
    work = store.rows.filter((r) => {
      if (r.status) return false; // already migrated / attempted
      if (progress.completed[r.doi]) return false; // checkpoint says done (Excel was locked earlier)
      if (CONFIG.category !== 'all' && r.category.toLowerCase() !== CONFIG.category) return false;
      return true;
    });
    if (CONFIG.limit > 0) work = work.slice(0, CONFIG.limit);
  }

  log(`Pending deposits this run: ${work.length}`);
  if (work.length === 0) {
    log('Nothing to do. Exiting.');
    return;
  }

  const handle = await openBrowser();
  const { page } = handle;
  page.setDefaultTimeout(CONFIG.selTimeout);
  page.setDefaultNavigationTimeout(CONFIG.navTimeout);

  let started = 0;
  let failed = 0;

  try {
    // --- One-time app setup ---
    log(`Navigating to ${WAT_URL}`);
    await page.goto(WAT_URL, { waitUntil: 'domcontentloaded', timeout: CONFIG.navTimeout });
    await settle(page, 1500);

    if (CONFIG.debugPause) {
      console.log('\n' + '-'.repeat(64));
      console.log(' DEBUG_PAUSE — Playwright Inspector is open.');
      console.log(' In the Inspector, click "Pick locator", then hover any control');
      console.log(' (side-panel button, Backstage tab, the dropdowns, the pubDOI box,');
      console.log(' a grid row, the Deposit menu item, the checkbox, etc.) and copy the');
      console.log(' locator it shows. You can also drive the app by hand to reach each step.');
      console.log(' Paste the locators back and they will be wired in. Press Resume to exit.');
      console.log('-'.repeat(64) + '\n');
      await page.pause();
    }

    await openBackstage(page);
    await selectComboOption(page, UI.envComboText, CONFIG.envOption); // Area → Live
    await selectComboOption(page, UI.typeComboText, PUBLICATION_LABEL); // Type → Digital Object Publications / Books
    await settle(page, 1000);
    log('App ready — beginning deposits.');

    // --- Per-DOI loop ---
    for (let i = 0; i < work.length; i++) {
      const row = work[i];
      const tag = `(${i + 1}/${work.length})`;
      log(`${tag} DOI ${row.doi} [${row.category || 'n/a'}]`);

      try {
        await depositOne(page, row.doi);

        const date = todayLabel();
        if (row.rowIdx >= 0) store.setStatus(row, CONFIG.statusValue, date);
        progress.completed[row.doi] = { status: CONFIG.statusValue, date, at: new Date().toISOString() };
        saveProgress(progress);
        started++;
      } catch (err) {
        if (err instanceof DryRunStop) {
          log(`${tag} dry-run reached the Deposit button successfully — stopping (no actual deposit).`);
          break;
        }
        failed++;
        const reason = (err as Error).message ?? String(err);
        warn(`${tag} FAILED ${row.doi}: ${reason}`);
        await screenshot(page, `fail-${row.doi}`);
        progress.failed[row.doi] = { reason, at: new Date().toISOString() };
        saveProgress(progress);
        if (CONFIG.markErrors && row.rowIdx >= 0) {
          store.setStatus(row, `error: ${reason.slice(0, 80)}`, todayLabel());
        }
        // Try to recover the UI for the next DOI (close any stray dialog).
        await page.keyboard.press('Escape').catch(() => {});
        await settle(page, 600);
      }

      if (i < work.length - 1) {
        await page.waitForTimeout(CONFIG.delayMs);
      }
    }
  } finally {
    await handle.close();
  }

  console.log('='.repeat(64));
  log(`Done. Started: ${started}  Failed: ${failed}  Total attempted: ${started + failed}`);
  if (failed > 0) log(`Failures logged to ${PROGRESS_FILE} and screenshots in ${SHOT_DIR}`);
}

main().catch((err) => {
  console.error('[Fatal]', err);
  process.exit(1);
});
