# IT3040 Assignment 1 — Option 1: Transliteration Testing - IT2321

Playwright-based automation that tests the Chat Sinhala transliteration function at
[https://www.pixelssuite.com/chat-translator](https://www.pixelssuite.com/chat-translator)
and records pass/fail results into the assignment Excel file.

The suite contains **50 negative** test cases (where the system produces an
incorrect transliteration) and **16 positive** test cases (where the system
produces the expected output) — 66 total.

## Prerequisites

- Python 3.11 or 3.12
- Google Chrome installed (the script uses your installed Chrome — no Chromium download required)

## Setup

Clone the repository and install dependencies:

```
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
pip install -U pip
pip install -r requirements.txt
```

You do **not** need to run `playwright install` — this project launches your
installed Chrome via `channel="chrome"`.

## Workflow

Edit `test_cases.tsv`, then run **one** command to load + test + record:

```
python load_test_cases.py --run
```

This loads your TSV into the Excel file, opens Chrome, runs every case
against the live site, writes Actual output + PASS/FAIL back to the Excel
file, and keeps the browser open until you press Ctrl+C.

> Always close `Assignment 1 - Test cases.xlsx` in Excel before running any
> script — Excel locks the file and overwrites changes.

## Step-by-step (run stages separately)

### 1. Edit test cases in `test_cases.tsv`

Tab-separated, one row per test case:

```
TC ID	Input length type	Input	Expected output
Neg_0001	S	meka Rs.1500n ganna puluwan	මේක Rs.1500න් ගන්න පුළුවන්
Neg_0002	M	USD 50 kiyanne kiyada LKR walin	USD 50 කියන්නේ කීයද LKR වලින්
```

Rules:
- **TC ID**: `Neg_XXXX` for cases where the system *fails* (50 required), `Pos_XXXX` for cases that pass (16 required).
- **Input length type**: `S` (≤ 30 chars), `M` (31–299), `L` (300–450).
- **Input**: the Singlish chat-style text.
- **Expected output**: the correct Sinhala transliteration.
- Separate fields with a TAB **or** 2+ spaces — the loader accepts either.

### 2. Load test cases into the Excel file

```
python load_test_cases.py
```

Reads `test_cases.tsv` and writes them into `Assignment 1 - Test cases.xlsx`,
extending the merged-cell template as needed. Old test data is cleared first.

To clear test data without loading new cases:
```
python load_test_cases.py --reset
```

### 3. Run the Playwright automation

```
python test_automation.py --excel "Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

A Chrome window opens, types each input into the translator, reads the Sinhala
output, and writes Actual output + Status (PASS/FAIL) back into the Excel file.
Press Ctrl+C in the terminal to stop when done.

For a faster, invisible run (e.g. while iterating on test cases):
```
python test_automation.py --excel "Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --headless --type-delay-ms 0 --slow-mo-ms 0 --save-every 1
```

### 4. Review and complete the Excel manually

Open `Assignment 1 - Test cases.xlsx` in Excel. The **Actual output** and
**Status** columns are filled in automatically. After reviewing, manually add
two more columns per assignment Appendix 2:
- **Singlish input types covered**
- **Evidence or rationale for the input type covered**

## Project files

| File | Purpose |
|---|---|
| `test_automation.py` | Playwright automation — runs your test cases against the live site |
| `load_test_cases.py` | Batch loader: reads `test_cases.tsv` → writes to Excel |
| `test_cases.tsv` | Test case data (50 negative + 16 positive) |
| `Assignment 1 - Test cases.xlsx` | Submission file (Appendix 2 template) |
| `requirements.txt` | Python dependencies |
| `.gitignore` | Files excluded from version control |

## Common issues

- **`PermissionError` when saving Excel**: file is open in Excel. Close it.
- **Data disappears after running script**: Excel was open during the run and overwrote the script's save. Close Excel first.
- **`ECONNRESET` from `playwright install`**: not needed — this project uses installed Chrome via `channel="chrome"`.

## License

This project is released under the MIT License — see [LICENSE](LICENSE).
