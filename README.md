# MPR Combiner (Usage Tab)

Combine the `Usage` / `Usage (Part B)` tabs from many MPR Excel files into **one combined Excel file**.

## Quick start

1) Put MPR files in `data/input/` (subfolders are OK).

2) Run:

```bash
./scripts/run.sh
```

Output:

- `data/output/combined-usage.xlsx`

## Next step (Excel Copilot prompt)

Use this prompt in Excel Copilot to format the combined sheet for upload:

```text
Format this Usage sheet for MPR upload: 1. Delete the 'COOP' column. 2. Update the 'Recipient Agency Number' column by combining the 'State' column, a dash '-', and the current RA number. Analyze the existing RA numbers in the sheet to learn the standard number length for this state, and pad the numbers with leading zeros to match that length before combining. 3. Keep only these columns in this exact order: Processor ID, Processor Name, Report Month, Report Year, State, Recipient Agency Number, Recipient Agency Name, Product Number, Product Name, USDA Material, EPDS DF, Case Qty, Used LBS.
```

## Notes

- **Sheet selection**: tries `Usage`, otherwise falls back to `Usage (Part B)`
- **Headers**: normalizes common header variations/broken headers
- **Need troubleshooting?** Run with `--diagnose true`

## Options

These are the available CLI switches (you can pass them to `./scripts/run.sh ...`).

- **`--input`** (required): input folder of Excel files  
  - Example: `--input "data/input"`
- **`--output`** (required): output file path  
  - Use `.xlsx` for Excel output, `.csv` for CSV output
- **`--recursive`** (default `false`): scan subfolders  
  - Note: `./scripts/run.sh` already sets `--recursive true`
- **`--sheet`** (default `Usage`): sheet name to extract  
  - If `Usage` isn’t found, it automatically falls back to `Usage (Part B)`
- **`--headerRow`** (default `1`): 1-based header row index in the sheet
- **`--includeSource`** (default `false`): include `__source_file`, `__source_sheet`, `__source_row` columns
- **`--diagnose`** (default `false`): print per-file extraction stats and skip reasons

