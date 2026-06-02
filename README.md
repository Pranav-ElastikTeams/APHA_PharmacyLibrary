# APHA NAPLEX QA Automation Script

Playwright-based QA script that sweeps NAPLEX Math Review question pages on `staging.pharmacylibrary.com`, verifies content against an Excel source of truth, and outputs a CSV report with per-page pass/fail details.

## What it checks

- Question stem is present and non-empty
- Reveal button (`input[name="choice"]`) exists
- Displayed answer matches expected answer from the QA Excel file
- Explanation block is present after reveal
- No broken equation or other images
- No raw MathML tags or HTML entities visible in rendered text
- No Unicode replacement characters (`�`)
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

## Usage

### Recommended: CDP mode (connect to your running Edge)

**Step 1 — Open the browser**

Double-click `open-edge.bat` (or run it from the terminal). It launches Edge with the remote debugging port open.

**Step 2 — Log in manually**

In the Edge window that opens, go to `pharmacylibrary.com` and log in. Keep the window open.

**Step 3 — Run the sweep**

```bash
npm run check
```

The script connects to that Edge window over CDP and opens a new tab for each question page.

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

## Debugging

Set `HEADLESS=false` in `.env` to watch the browser. Set `DEBUG_PAUSE=true` to open Playwright Inspector before the sweep starts.

If Cloudflare blocks the run, switch to CDP mode (the most reliable path — Edge already has a valid session) or refresh your cookies in `.env`.
