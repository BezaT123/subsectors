# Subsector Excel Extractor

Utilities for extracting structured JSON from Excel financial templates (`.xlsx`/`.xlsm`) using `extract.py`.

## Setup

- Requires Python 3.10+ and `openpyxl`.
- Install dependency:
  - `pip install openpyxl`

## Run a single file

- Command: `python extract.py --file "/path/to/your.xlsx"`
- What happens:
  - Loads the workbook and finds the `i_Setup`, `i_COS`, and optional `info` sheets (handles name variations like `i Setup`, `COS`, `Info`).
  - Builds a JSON file next to the Excel input with the same base name (e.g., `Financials.xlsm` -> `Financials.json`).
  - Prints a brief summary including counts and the top 5 products by weighting.

## Run a batch of company files

- Command: `python extract.py --batch "/path/to/directory"`
- Behaviour:
  - Scans the directory for non-hidden `.xlsm`/`.xlsx` files (skips `subsectors-example.xlsx`).
  - Processes each file individually and writes a per-file JSON alongside each Excel file.
  - Shows a success/failed count at the end.

## Output structure (per file)

- File: `<excel-basename>.json` written in the same directory as the source Excel.
- Top-level keys:
  - `extractedAt`: ISO timestamp
  - `sourceFile`: original Excel filename
  - `i_Setup`: field counts plus detailed field data and subtables
  - `i_COS`: product count plus `topProducts.byWeighting` (top 5)
  - `info`: benchmark/metadata if the sheet exists
