# AFS Footnote Compare (Excel VBA)

Compares the footnote text of Annual Financial Statements across two years for many
funds at once, inside one Excel workbook. Built for the case where each fund has one
PDF per year, the fund code (`XXXX-XX`) is in the filename, and note headings are bold.

Numbers are ignored when deciding whether a section changed. Only wording counts.

## Requirements

- Windows, Excel (macro-enabled workbook), and Word 2013 or later on the same machine.
  Word is used to open each PDF and read which paragraphs are bold. Adobe is not used.
- Two folders: one with the prior-year PDFs, one with the current-year PDFs.

## Setup

1. Create a new workbook and save it as `.xlsm`.
2. Alt+F11, File > Import File, pick `AFS_FootnoteCompare.bas`
   (or Insert > Module and paste the file contents).
3. Alt+F8, run `CompareFootnotes`.

## What happens on a run

1. First run asks for the 2024 and 2025 folders and stores them on a `Config` sheet.
2. It lists every fund code found in the filenames and asks how many to process now:
   `1, 5, 10, 25, 50, 100, 200, ALL`, or `0` to only rebuild the comparison.
3. Each fund's notes are extracted once into `Data_2024` / `Data_2025`: one row per fund,
   one column per bold heading, cell = that section's text. Later runs skip funds that
   already have a row. To redo a fund, delete its row and run again.
4. `Compare` is rebuilt every run: same heading columns, four rows per fund:
   2025 text, 2024 text, Same? (Yes/No per section), and Changed 2024>2025 with
   word-level markup `[-removed-] {+added+}` or `NEW IN 2025` / `REMOVED IN 2025`.
5. `Summary` has one row per fund with status, section counts, and which sections changed.
6. `Log` lists skipped files and extraction errors.

Word takes roughly 5 to 20 seconds per PDF, so start with a batch of 1 and check the
Data sheet before running ALL. Save the workbook between batches.

## Settings (top of the module)

| Constant | Default | Meaning |
|---|---|---|
| `YEAR_PRIOR` / `YEAR_CURRENT` | 2024 / 2025 | Labels and sheet names |
| `FUND_CODE_PATTERN` | `[A-Za-z0-9]{4}-[A-Za-z0-9]{2}` | Regex for the code in the filename |
| `NOTES_HEADER_PATTERN` | Notes to Financial Statements | Only pages with this line are read |
| `IGNORE_NUMBERS` | True | Digits become `#` for the Same? test |
| `SPLIT_RUNIN_HEADINGS` | True | A bold lead-in at the start of a paragraph is its own heading |
| `REPEAT_FURNITURE` | 3 | Bold text seen this many times (fund name, page header) is ignored |

## Known limits

- A section longer than about 32,000 characters is truncated in its cell (Excel limit).
- Two bold paragraphs in a row are merged into one heading, since headings sometimes wrap.
- If Word's PDF import loses bold on a particular PDF generator, headings will not be
  detected and everything lands under `(Intro)`. The Remark column shows the heading count.

This repo contains no data. It was drafted against fictional statements.
