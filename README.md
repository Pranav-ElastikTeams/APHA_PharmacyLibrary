# APHA NAPLEX QA Automation Script

Playwright-based QA script that sweeps NAPLEX Math Review question pages on `staging.pharmacylibrary.com`, verifies content against an Excel source of truth, and outputs a CSV report with per-page pass/fail details.

## What it checks

- Question stem is present and non-empty
- Reveal button (`input[name="choice"]`) exists
- Displayed answer matches expected answer from the QA Excel file
- Explanation block is present after reveal
- No broken equation or other images
- No raw MathML tags or HTML entities visible in rendered text
- No Unicode replacement characters (`?`)
- No JS console errors

Screenshots are saved for every page with issues.

## Prerequisites

- [Node.js](https://nodejs.org/) v18 or later
- Microsoft Edge installed (used by Playwright in CDP / persistent mode)

## Setup

```bash
npm install
npx playwright install msedge
```

Copy the environment template and fill in your credentials:

```bash
copy .env.example .env
```

Edit `.env` — see comments inside for each option.

Place the two Excel input files in the `data/` folder:

| File | Purpose |
|------|---------|
| `NAPLEX Math Review questions_29-4-26.xlsx` | DOI list (col A = S.No., col B = DOI) |
| `Naplex_Math-Review-Questions.xlsx` | QA data (col C = stem, col D = expected answer) |

## Key `.env` variables

| Variable | Default | Description |
|----------|---------|-------------|
| `BROWSER_MODE` | `cdp` | `cdp` = connect to running Edge via DevTools Protocol; `launch` = launch fresh Chromium |
| `BROWSERS` | `1` | Number of parallel browser instances. In CDP mode run `open-edge.bat <N>` first. |
| `CDP_BASE_PORT` | `9222` | First debug port. Script connects to this port, port+1, port+2, … for each browser. |
| `CONCURRENCY` | `20` | Total parallel worker tabs spread across all browser instances. |
| `HEADLESS` | `true` | Set to `false` to watch the browser. |
| `AUTH_METHOD` | `login` | `login` = fill the login form; `cookie` = inject `cf_clearance` + session cookie. |

## Usage

### Recommended: CDP mode (connect to your running Edge)

**Step 1 — Open browser windows**

Run `open-edge.bat` with the number of windows you want (default is 1):

```bat
open-edge.bat          :: opens 1 window on port 9222
open-edge.bat 3        :: opens 3 windows on ports 9222, 9223, 9224
```

On first run (no profile exists yet) the script opens a single fresh window for login, then exits. Log in to `pharmacylibrary.com`, close nothing, and re-run the same command — it will clone the authenticated session to the extra profiles automatically.

Set `BROWSERS` in `.env` to match the number you opened.

**Step 2 — Log in manually** (first run only)

In the Edge window that opens, go to `pharmacylibrary.com` and log in. Keep the window open, then re-run `open-edge.bat <N>`.

**Step 3 — Run the sweep**

```bash
npm run check
```

The script connects to those Edge windows over CDP and opens worker tabs for each question page.

### Alternative: Cookie injection mode

Get `cf_clearance` and `JSESSIONID` from your browser's DevTools (Application → Cookies on `pharmacylibrary.com`), add them to `.env`, set `AUTH_METHOD=cookie` and `BROWSER_MODE=launch`, then:

```bash
npm run check
```

## Output

| Path | Description |
|------|-------------|
| `output/qna_qa_report.csv` | Full per-page results — open in Excel or any CSV viewer |
| `screenshots/<doi>.png` | Screenshot for every page flagged with issues |

### CSV columns

`doi`, `question_no`, `url`, `http_status`, `has_question_text`, `page_stem`, `reveal_button_present`, `displayed_answer`, `expected_answer`, `answer_matches`, `broken_equation_images`, `broken_other_images`, `raw_mathml_visible`, `raw_entities_visible`, `explanation_present`, `missing_alt_images`, `console_errors`, `other_issues`, `notes`

## DOI finder — `npm run find-dois`

Scans the full production DOI catalogue and matches page text against a list of review questions to find the correct DOI for each.

### Input files (in `Data/Questions DOI extract/`)

| File | Purpose |
|------|---------|
| `mapped-question-doi.csv` | Review questions that need a DOI (`question_no`, `question_stem`, `doi`). Leave `doi` blank for questions to match; pre-fill it to skip them. |
| `NAPLEX only production_reporting dim_item 2026-05-27T1520.csv` | Full production DOI list to scan (4 600+ DOIs). |

### How to run

```bat
open-edge.bat 3        :: open N Edge windows and log in (first run only)
```

Set `BROWSERS=3` and `CONCURRENCY=20` (or higher) in `.env`, then:

```bash
npm run find-dois
```

The script visits every DOI page and compares the rendered question stem against each review question using a Dice-coefficient similarity (threshold 50%). Matches are printed as they are found.

### Resume support

Progress is checkpointed every 50 pages to `output/doi-scan-checkpoint.json`. If the run is interrupted, re-run `npm run find-dois` and it picks up from where it left off. Delete the checkpoint file before starting a brand-new full scan.

### Output

`output/may-2026-review-questions-with-dois.csv` — columns: `question_no`, `question_stem`, `doi`, `match_score`.

Successive runs accumulate candidates (existing rows are not overwritten). Pick the highest-scoring match per question.

## Deposit automation — `npm run deposit`

Deposits not-yet-migrated digital objects from the Atypon WAT console (`pharmacylibrary.com/wat`) to Impelsys, driven by the `DOs` tab of `Data/DOIs record/Migration_record.xlsx`. Fully independent of the QA and find-dois scripts; it reuses the **same Edge profile** so the login is shared.

### How it works

For each run it: clicks the **Administration** side-panel button → **Backstage** → selects the environment (**Live**) and publication type (**Digital Object Publications** by default), then for every pending DOI it types the DOI into the **DOI** search box, right-clicks the matching grid row → **Deposit**, ticks **Impelsys Migration Feed**, clicks **Deposit**, clicks the **OK/Yes** confirmation if one appears, waits for the "deposit started" notification, marks the row **Status = `started`** with today's date in the Excel file, closes the notification, waits `DEPOSIT_DELAY_MS`, and moves on.

### Which rows are processed

- A row is **pending** when its `Status from Atypon` column is blank.
- The `Category` column is filtered by `DEPOSIT_CATEGORY` (`NAPLEX`, `other`, or `all`).
- Because the Status column is written back as each deposit starts, **re-running resumes automatically** — already-started rows are skipped.

### Run it

```bat
open-edge.bat 1        :: open Edge (first run: log in to pharmacylibrary.com, then re-run)
```

Then, with the `DEPOSIT_*` values set in `.env`:

```bash
npm run deposit
```

### Key `.env` variables

| Variable | Default | Description |
|----------|---------|-------------|
| `DEPOSIT_CATEGORY` | `NAPLEX` | Which `Category` rows to deposit: `NAPLEX` / `other` / `all` |
| `DEPOSIT_PUBLICATION_TYPE` | `digital` | `digital` = Digital Object Publications, `books` = Books |
| `DEPOSIT_ENV_OPTION` | `Live` | Environment dropdown option on the Backstage screen |
| `DEPOSIT_SINGLE_DOI` | _(blank)_ | Deposit exactly one DOI and stop |
| `DEPOSIT_DELAY_MS` | `5000` | Delay between two deposits |
| `DEPOSIT_LIMIT` | `0` | Cap the number of deposits this run (`0` = no limit) |
| `DEPOSIT_DRY_RUN` | `false` | Do everything except the final Deposit/confirm — verifies the flow safely |
| `DEPOSIT_MARK_ERRORS` | `false` | Write `error: …` to Status on failure so it is skipped next run |

### First-time verification

The Atypon UI is a Vaadin app with heavy shadow DOM; the controls are targeted by their
accessible role/name (which pierces shadow DOM). Run once with `DEPOSIT_DRY_RUN=true` and
`DEPOSIT_LIMIT=1` to walk the whole flow up to — but not including — the actual deposit.
Failures are screenshotted to `output/deposit-screenshots/` and progress is mirrored to
`output/deposit-progress.json`.

The fixed UI control names/labels (side-panel button, combo-box labels, option text, etc.)
are **hard-coded** in `src/deposit.ts` (the `UI` object). If the UI ever changes, edit them
there. To re-pick a locator, set `DEBUG_PAUSE=true` (and optionally `$env:PWDEBUG=1`) to drop
into the Playwright Inspector right after `/wat` loads.

> Note: saving Status back rewrites the workbook via SheetJS, which preserves all data on both sheets but does not retain cell styling. Keep the file closed while the script runs (an open file in Excel locks writes; the script keeps a JSON checkpoint as backup if that happens).

## Debugging

Set `HEADLESS=false` in `.env` to watch the browser. Set `DEBUG_PAUSE=true` to open Playwright Inspector before the sweep starts.

If Cloudflare blocks the run, switch to CDP mode (the most reliable path — Edge already has a valid session) or refresh your cookies in `.env`.
