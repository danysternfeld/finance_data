import os
import sys
import xml.etree.ElementTree as ET

from pandas import DataFrame

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from get_paper_data import parse_bizportal_fund


# Parse an Excel file in XML format (xls) and extract the full paper data.
def read_holdings_excel(path: str) -> DataFrame:
    with open(path, "r", encoding="utf-8") as workbook:
        text = workbook.read().lstrip()
    root = ET.fromstring(text)
    ns = {
        "ss": "urn:schemas-microsoft-com:office:spreadsheet",
        "o": "urn:schemas-microsoft-com:office:office",
        "x": "urn:schemas-microsoft-com:office:excel",
    }
    ws = root.find("ss:Worksheet", ns)
    if ws is None:
        raise ValueError("The XML workbook does not contain a worksheet.")
    tbl = ws.find("ss:Table", ns)
    if tbl is None:
        raise ValueError("The XML worksheet does not contain a table.")

    rows = []
    header = None
    started = False

    for row in tbl.findall("ss:Row", ns):
        vals = []
        for cell in row.findall("ss:Cell", ns):
            data = cell.find("ss:Data", ns)
            vals.append(data.text if data is not None else "")

        if not started:
            if vals and any(v == "מספר נייר" for v in vals):
                header = [str(v).strip() for v in vals]
                started = True
            continue

        if not vals or all((v is None or str(v).strip() == "") for v in vals):
            continue

        if vals and vals[0] and vals[0].isdigit():
            rows.append([str(v) for v in vals])
        else:
            break

    if header is None:
        return DataFrame()

    max_len = max(len(header), max((len(r) for r in rows), default=0))
    normalized = []
    for row in rows:
        padded = row[:max_len] + [""] * (max_len - len(row))
        normalized.append(padded)

    df = DataFrame(normalized, columns=header[:max_len])
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    return df


def enrich_holdings_with_bizportal(holdings: DataFrame) -> DataFrame:
    """Return a copy of holdings with Bizportal data added for each paper."""
    paper_id_column = "מספר נייר"
    if paper_id_column not in holdings.columns:
        raise ValueError(f"Holdings dataframe is missing the {paper_id_column!r} column.")

    enriched_holdings = holdings.copy()
    enrichment = []

    for paper_id in enriched_holdings[paper_id_column]:
        paper_id = str(paper_id).strip()
        if paper_id and paper_id.isdigit():
            fund = parse_bizportal_fund(paper_id)
            enrichment.append({
                "fund_name": fund.get("fund_name"),
                "management_fee": fund.get("management_fee"),
                "three_year_return_percent": fund.get("three_year_return_percent"),
                "sharpe_ratio_12_months": fund.get("sharpe_ratio_12_months"),
                "bizportal_url": fund.get("url"),
            })
        else:
            enrichment.append({
                "fund_name": None,
                "management_fee": None,
                "three_year_return_percent": None,
                "sharpe_ratio_12_months": None,
                "bizportal_url": None,
            })

    enrichment_df = DataFrame(
        enrichment,
        index=enriched_holdings.index,
        columns=[
            "fund_name",
            "management_fee",
            "three_year_return_percent",
            "sharpe_ratio_12_months",
            "bizportal_url",
        ],
    )

    for column in enrichment_df.columns:
        enriched_holdings[column] = enrichment_df[column]
    return enriched_holdings


def read_holdings_excel_and_enrich(path: str) -> DataFrame:
    """Read holdings from an Excel file and enrich them with Bizportal data."""
    holdings = read_holdings_excel(path)
    enriched_holdings = enrich_holdings_with_bizportal(holdings)
    return enriched_holdings

def fix_percentage_columns(holdings: DataFrame) -> DataFrame:
    """Convert percentage columns to float and multiply by 100 to get actual percentage values."""
    percentage_columns = ["%  מהתיק", "% שינוי יומי", "רווח ב-%", "תשואה 12 חודשים"]
    for pct_col in percentage_columns:
        if pct_col in holdings.columns:
            holdings[pct_col] = holdings[pct_col].astype(float) * 100
    # remove the % sign from management_fee column if it exists, convert to float 
    if "management_fee" in holdings.columns:
        try:
            holdings["management_fee"] = holdings["management_fee"].str.replace("%", "").astype(float) 
        except ValueError:
            print(f"Warning: Could not convert management_fee to float. Please check the data format. {holdings['management_fee']}")
            pass
    holdings["פרופיל חשיפה"] = holdings["פרופיל חשיפה"].str.replace(r'^$|None|00', '\'00', regex=True)
    return holdings

if __name__ == "__main__":
    holdings = read_holdings_excel_and_enrich(r"c:\Users\danys\OneDrive\Documents\scripts\finance_data\אחזקות.xls")
    print(holdings.head())
    print(holdings.columns.tolist())
    print(holdings)
    holdings = fix_percentage_columns(holdings)

    holdings.to_excel(r"c:\Users\danys\OneDrive\Documents\scripts\finance_data\תחבצ.xlsx", index=False)   


