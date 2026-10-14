# finance_data

`finance_data` is a Python-based portfolio analysis toolkit for Israeli investment data. It reads local holdings exports, enriches them with Bizportal metadata, combines them with IBKR and foreign-security data, and produces a prepared Excel workbook for portfolio/performance analysis.

It is designed around a workflow where a holdings export (for example from a local Excel/XML fund holdings report) is combined with market and fund data to build a consolidated analysis sheet.

## What this project does

- Reads holdings data from an exported Excel/XML workbook.
- Extracts paper identifiers and portfolio columns from the holdings file.
- Enriches each paper with Bizportal fund metadata, including:
  - fund name
  - management fee
  - 3-year return
  - Sharpe ratio
  - Bizportal URL
- Pulls additional data for IBKR positions and foreign securities.
- Generates an Excel workbook based on a template, filling a report with the merged data.

## Repository structure

- `get_paper_data.py` – Fetches and parses Bizportal fund pages and extracts fund metadata.
- `read_holdings_excel.py` – Reads holdings exports, normalizes the table, and enriches holdings with Bizportal data.
- `ibkr_papers.py` – Reads IBKR positions and adds performance and Sharpe metrics using Yahoo Finance mappings.
- `foreign_papers.py` – Utilities for foreign paper and market data calculations.
- `make_performance_xl.py` – Builds the final performance Excel workbook from the holdings and IBKR data.
- `make_version_file.py` – Generates version metadata used by the packaged tool.
- `build.bat` – Windows packaging script that creates a bundled executable and zipped output.
- `.gitignore` – Standard project ignore rules.

## Typical workflow

1. Prepare a holdings export file (for example an `אחזקות.xls`-style workbook).
2. Place the relevant files in the expected folder structure used by the scripts.
3. Run the workbook generation flow.
4. Generate a merged Excel file for portfolio analysis.

The project is strongly oriented toward a local Windows desktop workflow and uses a template Excel file plus additional data files such as `Open_Positions.csv` and ticker mappings.

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

This is a local operational script rather than a library package, so usage depends on the expected file layout and the included Excel template.

## Notes

- The project works with Israeli finance data and references Bizportal heavily.
- It is optimized for a Windows-based workflow and includes a batch build helper.
- Some scripts are written as operational utilities, so they may expect specific file names and directory conventions.

## License

This repository does not currently include a license file. If you plan to distribute or reuse the project publicly, you may want to add an explicit license such as MIT or Apache 2.0.

## Maintainer

This repository appears to be a personal or internal analytics utility. Use with the understanding that it is tailored to a specific reporting workflow and data source layout.
