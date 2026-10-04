from io import StringIO
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup


def parse_bizportal_fund(paperID: str, timeout: float = 20) -> dict[str, object]:
    """Fetch a Bizportal traded-fund page and parse its metadata and tables.

    ``tables`` is a list of dictionaries containing a table name and its
    contents as a pandas DataFrame. Performance metrics include the
    three-year return from the performance page and Sharpe ratio from the
    main page, preserving the exact text Bizportal displays.
    """
    url = f"https://www.bizportal.co.il/tradedfund/quote/generalview/{paperID}"
    parsed_url = urlparse(url)
    hostname = (parsed_url.hostname or "").lower()
    if (
        parsed_url.scheme not in {"http", "https"}
        or not (hostname == "bizportal.co.il" or hostname.endswith(".bizportal.co.il"))
    ):
        raise ValueError("url must be an HTTP(S) Bizportal URL")
    if timeout <= 0:
        raise ValueError("timeout must be greater than zero")

    headers = {"User-Agent": "Mozilla/5.0 (compatible; fund-data-parser/1.0)"}
    response = requests.get(url, headers=headers, timeout=timeout)
    response.raise_for_status()

    soup = BeautifulSoup(response.content, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    meta_title = soup.find("meta", attrs={"name": "title"})
    description = soup.find("meta", attrs={"name": "description"})
    management_fee_label = next(
        (
            label
            for label in soup.find_all("dt")
            if label.get_text(" ", strip=True) == "דמי ניהול"
        ),
        None,
    )
    management_fee_value = None
    if management_fee_label:
        management_fee_element = management_fee_label.find_next_sibling("dd")
        if management_fee_element:
            management_fee_value = management_fee_element.get_text(" ", strip=True)
    sharpe_label = next(
        (
            label
            for label in soup.select("li span.label")
            if label.get_text(" ", strip=True).startswith("שארפ")
        ),
        None,
    )
    sharpe_value = (
        sharpe_label.find_next_sibling("span", class_="num").get_text(strip=True)
        if sharpe_label and sharpe_label.find_next_sibling("span", class_="num")
        else None
    )

    tables = []
    used_names = set()
    for index, table in enumerate(soup.find_all("table"), start=1):
        caption = table.find("caption")
        heading = table.find_previous(["h1", "h2", "h3"])
        if caption:
            name = caption.get_text(" ", strip=True)
        elif heading:
            name = heading.get_text(" ", strip=True)
        else:
            name = f"table_{index}"

        base_name = name
        suffix = 2
        while name in used_names:
            name = f"{base_name}_{suffix}"
            suffix += 1
        used_names.add(name)

        frames = pd.read_html(StringIO(str(table)))
        if frames and not frames[0].empty:
            tables.append({"name": name, "data": frames[0]})

    performance_url = url.replace("/quote/generalview/", "/quote/performance/")
    if performance_url == url:
        raise ValueError("url must point to a Bizportal fund generalview page")
    performance_response = requests.get(performance_url, headers=headers, timeout=timeout)
    performance_response.raise_for_status()

    metrics = {
        "three_year_return_percent": None,
        "sharpe_ratio_12_months": sharpe_value,
    }
    for frame in pd.read_html(StringIO(performance_response.text)):
        if frame.empty or len(frame.columns) < 2:
            continue
        labels = frame.iloc[:, 0].astype(str).str.strip()
        for row_index, label in labels.items():
            values = frame.iloc[row_index, 1:]
            if label == "3 שנים":
                metrics["three_year_return_percent"] = pd.to_numeric(
                    str(values.iloc[0]).replace("%", "").replace(",", ""),
                    errors="coerce",
                )

    return {
        "url": url,
        "title": title,
        "fund_name": meta_title.get("content", "").strip() if meta_title else title,
        "description": description.get("content", "").strip() if description else "",
        "management_fee": management_fee_value,
        "tables": tables,
        **metrics,
    }

"""
fund = parse_bizportal_fund(
    1170703
)
print(fund["fund_name"])
print(f"3-year return: {fund['three_year_return_percent']}%")
print(f"Sharpe ratio (12 months, from main page): {fund['sharpe_ratio_12_months']}")
print(f"Parsed {len(fund['tables'])} non-empty overview tables")
"""