#!/usr/bin/env python3
import os
import argparse
import time
import random
import logging
from datetime import datetime

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import pandas as pd

# base directory and data directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(message)s",
    level=logging.INFO
)


def fetch_us_sales():
    HEADERS = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/115.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://www.goodcarbadcar.net/"
    }

    url = "https://www.goodcarbadcar.net/2025-us-auto-sales-figures-by-brand-brand-rankings/"

    session = requests.Session()
    session.headers.update(HEADERS)
    retries = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    time.sleep(random.uniform(1.0, 3.0))

    resp = session.get(url, timeout=10)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "lxml")
    table = soup.find("table", id="table_6")
    if table is None:
        raise RuntimeError("🚨 could not find table_6")

    rows = table.find_all("tr", attrs={"data-row-index": True})
    data = [
        [td.get_text(strip=True).replace(",", "") for td in tr.find_all("td")]
        for tr in rows
    ]
    cols = ["Brand","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    df = pd.DataFrame(data, columns=cols)
    for m in cols[1:]:
        df[m] = df[m].astype(int)

    return df


def fetch_current_year_sales(year: int = 2025):
    df = fetch_us_sales()
    month_map = {
        m: pd.to_datetime(f"{year}-{i:02d}-01")
        for i, m in enumerate(
            ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
            start=1
        )
    }
    return df.rename(columns=month_map)


def update_2025_sales_only(excel_file_arg: str, sheet_name: str = "Brands"):
    # resolve path
    if os.path.isabs(excel_file_arg):
        excel_path = excel_file_arg
    else:
        excel_path = os.path.join(DATA_DIR, excel_file_arg)

    existing = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", dtype=str)
    existing = existing.rename(columns={existing.columns[0]: "Automaker", existing.columns[1]: "Brand"})
    existing.columns = [
        pd.to_datetime(c) if isinstance(c, datetime) else c
        for c in existing.columns
    ]
    existing.set_index("Brand", inplace=True)

    new_df = fetch_current_year_sales().set_index("Brand")
    for month_ts in new_df.columns:
        existing[month_ts] = new_df[month_ts].astype(int)

    existing = existing.reset_index()
    cols_order = ["Automaker","Brand"] + [c for c in existing.columns if isinstance(c, pd.Timestamp)]
    existing = existing[cols_order]

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        existing.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    parser = argparse.ArgumentParser(description="Excel 파일의 'Brands' 시트를 2025년 데이터로 덮어씌웁니다.")
    parser.add_argument("excel_file", help="data 폴더 내 파일명 또는 절대경로")
    parser.add_argument("-s", "--sheet", default="Brands", help="대상 시트 이름 (기본: Brands)")
    args = parser.parse_args()

    try:
        update_2025_sales_only(args.excel_file, args.sheet)
        print(f"✅ '{args.excel_file}' 업데이트 완료")
    except Exception as e:
        print(f"❌ 업데이트 실패: {e}")

if __name__ == "__main__":
    main()