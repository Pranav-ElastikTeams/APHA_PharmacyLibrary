# Migration plan — DOs then Books, by 31 July 2026

## The numbers

| | Pending DOIs |
|---|---|
| DOs tab — NAPLEX | 4,656 |
| DOs tab — other | 3,469 |
| Books tab (all 11 tags) | 1,161 |
| **Total** | **9,286** |

Nothing has been deposited yet — every row on both tabs has a blank Status.

**Throughput:** ~15 s of UI work per deposit + `DEPOSIT_DELAY_MS`. At the 5,000 ms
delay this plan assumes, that is **20 s/item**, so:

- 300-DOI batch ≈ **1 hr 40 min** raw, ~1 hr 55 min with retries/lag
- Whole job ≈ 51.6 hrs raw, **~59 hrs** with a 15% buffer

## The scheduling problem — read this first

11 working days remain (Fri 17 Jul, then Mon–Fri 20–24 and 27–31). At 4 hrs/day that is
**44 attended hours against a ~59-hour job**. It does not fit, at any safe delay setting.

The fix is that **the script does not need you watching it.** It needs a logged-in Edge
session over CDP and a machine that stays awake. So the plan is **three batches of 300 per
day**: two while you are working, one launched as you leave that finishes on its own.

- 3 × 300 = 900 DOIs/day ≈ 5 hrs 45 min of wall-clock machine time
- Attended: ~3 hr 50 min (batches 1 and 2) — fits your 4 hours
- Unattended: ~1 hr 55 min (batch 3) — runs after you stop

If leaving it running unattended is not an option, **the 31st is not reachable** and the
call to make now is which scope slips — the honest choice is to finish all 8,125 DOs by the
31st and let Books (1,161) land in the first days of August.

## Order of work

Batching needs no bookkeeping from you. The script skips any row with a non-blank Status,
so **each batch is the same command re-run** — it picks up the next 300 pending rows in the
selected category automatically. When a category is exhausted it prints
`Nothing to do. Exiting.` — that is your signal to move to the next phase.

1. **Phase 1 — DOs / NAPLEX** (4,656) → ~16 batches
2. **Phase 2 — DOs / other** (3,469) → ~12 batches
3. **Phase 3 — Books** (1,161) → ~4 batches

Books must be a separate run: `deposit.ts:544` refuses to start when the sheet and the
publication-type dropdown disagree.

## Day-by-day target

| Day | Date | DOIs | Cumulative | Phase |
|---|---|---|---|---|
| 1 | Fri 17 Jul | 300 | 300 | Setup + smoke + 1 batch |
| 2 | Mon 20 Jul | 900 | 1,200 | NAPLEX |
| 3 | Tue 21 Jul | 900 | 2,100 | NAPLEX |
| 4 | Wed 22 Jul | 900 | 3,000 | NAPLEX |
| 5 | Thu 23 Jul | 900 | 3,900 | NAPLEX |
| 6 | Fri 24 Jul | 900 | 4,800 | NAPLEX ends (4,656) → switch to `other` |
| 7 | Mon 27 Jul | 900 | 5,700 | other |
| 8 | Tue 28 Jul | 900 | 6,600 | other |
| 9 | Wed 29 Jul | 900 | 7,500 | other |
| 10 | Thu 30 Jul | 900 | 8,400 | DOs end (8,125) → switch to Books |
| 11 | Fri 31 Jul | 886 | 9,286 | Books complete |

There is **no slack in this schedule**. Any day you can bank extra — a longer evening run,
a weekend batch — buy it back early rather than counting on a clean final week.

## What to change in `.env`

### Once, before you start

```ini
DEPOSIT_DELAY_MS=5000        # currently 10000 — halving it saves ~13 hrs overall
DEPOSIT_LIMIT=300            # currently 2 — this is the batch size
DEPOSIT_DRY_RUN=false
DEPOSIT_MARK_ERRORS=true     # already set; keeps failures from blocking the next batch
```

`DEPOSIT_DELAY_MS=10000` is the single biggest cost in the current config. **Only drop it to
5,000 if you know why it was raised** — if it was to stop the submissions page falling
behind, leave it at 10,000 and accept that the 31st needs a 4th batch each day.

Batch size can be 200–400 to taste. 300 is the recommendation: 200 means more babysitting
(~46 batches), 400 pushes a single batch past two hours.

| `DEPOSIT_LIMIT` | Batch runtime (@20 s/item) |
|---|---|
| 200 | ~1 hr 7 min |
| 300 | ~1 hr 40 min |
| 400 | ~2 hr 13 min |

### Phase 1 — DOs / NAPLEX (days 1–6)

```ini
DEPOSIT_SHEET=DOs
DEPOSIT_PUBLICATION_TYPE=digital
DEPOSIT_CATEGORY=NAPLEX
DEPOSIT_LIMIT=300
```

### Phase 2 — DOs / other (days 6–10)

Change **one line** when phase 1 prints `Nothing to do`:

```ini
DEPOSIT_CATEGORY=other
```

### Phase 3 — Books (days 10–11)

Change **three lines** together — the script hard-fails if these disagree:

```ini
DEPOSIT_SHEET=Books
DEPOSIT_PUBLICATION_TYPE=books
DEPOSIT_CATEGORY=all
```

`all` covers all 1,161 rows. To split by tag instead, the Books tags are: MediaLibrary 457,
PfirstModules 138, PfirstCases 127, PfirstFaculty 126, Books 72, OtcDecisionTrees 62,
OtcTopic 52, OtcCases 47, OtcFaculty 47, CulturalToolkit 33, TechnicianSeries 6.

## Day 1 (today) — validate before committing to the pace

The 20 s/item figure is inherited from the earlier estimate, not measured. Today's job is to
prove it, because every date above depends on it.

1. `open-edge.bat 1`, log in to pharmacylibrary.com by hand.
2. Dry run: `DEPOSIT_DRY_RUN=true`, `DEPOSIT_LIMIT=2` → `npm run deposit`. Confirms the
   selectors still match the live UI without depositing anything.
3. Live smoke: `DEPOSIT_DRY_RUN=false`, `DEPOSIT_LIMIT=2` → check both rows land on the
   submissions page and both got `started` + today's date in the sheet.
4. First real batch: `DEPOSIT_LIMIT=300`. **Time it.** If it runs materially longer than
   ~1 hr 40 min, the whole schedule above needs redoing — flag it before Monday.

## Running rules

- **Close Migration_record.xlsx before every batch.** With the file open in Excel the write
  fails, and progress survives only in `output/deposit-progress.json`.
- **Check the submissions page** after each batch — `started` means the deposit was accepted,
  not that it succeeded.
- **Failures don't block you.** With `DEPOSIT_MARK_ERRORS=true` a failed DOI gets
  `error: ...` in Status and is skipped next run. Screenshots land in
  `output/deposit-screenshots/`, reasons in `output/deposit-progress.json`. Sweep them at the
  end by clearing those Status cells and re-running.
- **The evening batch may die on session expiry** — `cf_clearance` and `JSESSIONID` don't
  live forever, and an unattended run has nobody to re-login. The script now detects this
  (HTTP 401) and stops immediately with a `SESSION EXPIRED — ... log in again` banner and
  **exit code 2**, rather than timing out on a missing "Administration" button. DOIs not yet
  attempted are left blank so the next run retries them — nothing is poisoned. Check the log
  each morning: a short overnight batch is ground to make up, not a one-off.
- **Disable sleep/hibernate** on the machine, or the unattended batch stops when the screen does.
- `CONCURRENCY=20` in `.env` does **not** apply here — `deposit.ts` is single-browser and
  sequential. There is no parallelism to turn on without reworking the script.
