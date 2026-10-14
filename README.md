# finance_data

`finance_data` is a Python-based portfolio analysis toolkit focused on Israeli and international investment data. It reads local holdings exports, enriches Israeli funds with Bizportal metadata, combines them with IBKR and foreign-security data (via Yahoo Finance), and produces a prepared Excel workbook for portfolio/performance analysis.

It is designed around a workflow where a holdings export (for example from a local Excel/XML fund holdings report) is combined with market and fund data to build a consolidated analysis sheet.

## What this project does

- Reads holdings data from an exported Excel/XML workbook.
- Extracts paper identifiers and portfolio columns from the holdings file.
- Enriches Israeli traded funds with Bizportal fund metadata, including:
  - fund name
  - management fee
  - 3-year return
  - Sharpe ratio
  - Bizportal URL
- Pulls additional data for IBKR positions and foreign securities using Yahoo Finance (via the `yfinance` library) and local ticker mappings.
- Generates an Excel workbook based on a template, filling a report with the merged data.

## International / non-Israeli data

The project supports non-Israeli (foreign) securities and uses Yahoo Finance to fetch performance and metrics for those tickers. Important details:

- `ibkr_papers.py` reads IBKR Open Positions (`Open_Positions.csv`) and uses a local mapping file (expected: `excel_files/TICKER in YF.xlsx`) to translate IBKR/local tickers to Yahoo Finance tickers.
- `yfinance` is used to compute performance and Sharpe ratios for foreign securities.
- The workflow performs currency conversion for non-ILS positions (exchange rates fetched from an external API) so values can be represented in ILS inside the final report.

If you add support for more markets or change ticker mapping conventions, update `excel_files/TICKER in YF.xlsx` to ensure correct Yahoo tickers are queried.

## Repository structure

- `get_paper_data.py` – Fetches and parses Bizportal fund pages and extracts fund metadata.
- `read_holdings_excel.py` – Reads holdings exports (XML-format xls), normalizes the table, and enriches holdings with Bizportal data.
- `ibkr_papers.py` – Reads IBKR positions and adds performance and Sharpe metrics using Yahoo Finance mappings.
- `foreign_papers.py` – Utilities for foreign paper and market data calculations.
- `make_performance_xl.py` – Builds the final performance Excel workbook from the holdings and IBKR data.
- `make_version_file.py` – Generates version metadata used by the packaged tool.
- `build.bat` – Windows packaging script that creates a bundled executable and zipped output.
- `.gitignore` – Standard project ignore rules.
- `excel_files/` (expected) – Template files and mapping spreadsheets (e.g., `Analysis_template.xlsx`, `TICKER in YF.xlsx`, composition sheets).

## Typical workflow

1. Prepare a holdings export file (for example an `אחזקות.xls`-style workbook exported as XML).
2. Place the relevant files in the expected folder structure used by the scripts (see `excel_files/`).
3. Run the workbook generation flow (see example usage below).
4. The scripts will create a merged Excel file for portfolio analysis based on the template.

The project is oriented toward a local Windows desktop workflow and uses a template Excel file plus additional data files such as `Open_Positions.csv` and ticker mappings.

## Dependencies

The scripts rely on the following Python packages:

- pandas
- requests
- beautifulsoup4
- openpyxl
- yfinance
- pyinstaller (for packaging)

Install them with:

```bash
pip install pandas requests beautifulsoup4 openpyxl yfinance pyinstaller
```

## Example usage

The main workbook-generation script is intended to run against the local project directory and a prepared folder structure:

```bash
python make_performance_xl.py "C:\path\to\input_folder"
```

There is also a Windows helper `build.bat` that runs `pyinstaller` to package the main script into a single executable and creates a zip archive of outputs.

## Notes & caveats

- The project handles both Israeli and international data; Bizportal calls are used for Israeli traded funds while `yfinance` is used for foreign tickers.
- Many scripts expect specific file names and directory conventions (for example `אחזקות.xls`, `Open_Positions.csv`, Excel templates under `excel_files/`).
- The code mixes English and Hebrew column names and is tailored to the author's internal reporting flow. You may need to adapt file paths and filenames to your environment.
- Network calls (Bizportal, exchange-rate API, Yahoo Finance) are synchronous and may need error handling, retries, or rate-limit handling for production use.

## License

This repository does not currently include an explicit license file. If you plan to distribute or reuse the project publicly, consider adding a license (for example MIT or Apache-2.0).

## Maintainer / Contact

This repository appears to be a personal or internal analytics utility. Use with the understanding that it is tailored to a specific reporting workflow and data source layout.
