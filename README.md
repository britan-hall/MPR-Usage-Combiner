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

> Format this Usage sheet for MPR upload: 1. Delete the 'COOP' column. 2. Update the 'Recipient Agency Number' column by combining the 'State' column, a dash '-', and the current RA number. Analyze the existing RA numbers in the sheet to learn the standard number length for this state, and pad the numbers with leading zeros to match that length before combining. 3. Keep only these columns in this exact order: Processor ID, Processor Name, Report Month, Report Year, State, Recipient Agency Number, Recipient Agency Name, Product Number, Product Name, USDA Material, EPDS DF, Case Qty, Used LBS.

## Notes

- **Sheet selection**: tries `Usage`, otherwise falls back to `Usage (Part B)`
- **Headers**: normalizes common header variations/broken headers
- **Need troubleshooting?** Run with `--diagnose true`

