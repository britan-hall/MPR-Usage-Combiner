# MPR Combiner (Usage Tab)

Combine the `Usage` or `Usage (Part B)` sheets from many MPR Excel files into **one combined workbook**.

## Quick start

1. Create the data folders:

```bash
mkdir -p data/input data/output
```

2. Copy your MPR Excel files into `data/input/` (subfolders are fine).

3. Run:

```bash
./scripts/run.sh
```

**Output** (when you don’t pass `--output`):

- Writes to `data/output/` using: `99_<STATE>_<QUARTER>.xlsx`  
  Example: `99_AL_Q4.xlsx`
- **State**: taken from the `State` column; if more than one state appears, the name uses `MULTI`; if none, `UNKNOWN`.
- **Quarter**: inferred from `Report Month` when present; otherwise from filenames like `...-2025-10-...`. Pass `--output path/to/file.xlsx` to choose the path yourself.

## Next step — Excel Copilot

Paste the block below into Excel Copilot to clean the combined sheet for upload:

```text
1. Delete rows that are empty or nearly empty (no meaningful data in most cells).

2. Move rows with a missing Processor ID to a new sheet named "Missing Processor ID". Leave only rows with a valid Processor ID on the main Usage sheet.

3. Remove the COOP column if it exists.

4. Set Recipient Agency Number to: State + "-" + RA number (digits only). For each state, look at existing RA numbers in the sheet, infer the usual digit length for that state, and left-pad the RA with zeros to that length before combining with State and the dash.

5. Strip thousands separators: remove comma characters from numbers (e.g. "1,234" → 1234).

6. Treat parentheses as negative numbers: replace accounting-style "(123)" with -123.

7. Normalize month names to numbers where needed (e.g. "December" → 12).

8. Keep only these columns, in this exact order:
Processor ID, Processor Name, Report Month, Report Year, State, Recipient Agency Number, Recipient Agency Name, Product Number, Product Name, USDA Material, EPDS DF, Case Qty, Used LBS.
```

## Notes

- **Sheet selection**: uses `Usage` if present; otherwise `Usage (Part B)`.
- **Headers**: common typos and broken spacing in column names are normalized.
- **Troubleshooting**: run with `--diagnose true` to see per-file extraction details.

## Options

Pass these to `./scripts/run.sh` or to `java -jar ...` after `--`.

| Switch | Meaning |
|--------|---------|
| `--input` | **Required** unless you rely on the script default (`data/input`). Folder containing `.xlsx` / `.xlsm` / `.xls` files. |
| `--output` | **Optional.** Output file (`.xlsx` or `.csv`). If omitted, see **Output** above. |
| `--recursive` | `true` \| `false` — include subfolders. The script sets `true` by default. |
| `--sheet` | Sheet name to read (default `Usage`; fallback to `Usage (Part B)` when appropriate). |
| `--headerRow` | 1-based row index of the header row (default `1`). |
| `--includeSource` | `true` to add `__source_file`, `__source_sheet`, `__source_row` columns. |
| `--diagnose` | `true` to print extraction stats and skip reasons to stderr. |
