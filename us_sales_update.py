#!/usr/bin/env python3
import argparse
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import pandas as pd

def fetch_us_sales():
    url = "https://www.goodcarbadcar.net/2025-us-auto-sales-figures-by-brand-brand-rankings/"
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")
    table = soup.find("table", id="table_6")
    if table is None:
        raise RuntimeError("🚨 could not find table_6")

    rows = table.find_all("tr", attrs={"data-row-index": True})
    data = [[td.get_text(strip=True).replace(",", "") for td in tr.find_all("td")] for tr in rows]
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

def update_2025_sales_only(excel_path: str, sheet_name: str = "Brands"):
    existing = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", dtype=str)
    existing = existing.rename(columns={existing.columns[0]:"Automaker", existing.columns[1]:"Brand"})
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
    parser.add_argument("excel_file", help="업데이트할 Excel 파일 경로")
    parser.add_argument("-s", "--sheet", default="Brands", help="대상 시트 이름 (기본: Brands)")
    args = parser.parse_args()

    try:
        update_2025_sales_only(args.excel_file, args.sheet)
        print(f"✅ '{args.excel_file}' 의 '{args.sheet}' 시트를 2025년 데이터로 업데이트했습니다.")
    except Exception as e:
        print(f"❌ 업데이트 실패: {e}")

if __name__ == "__main__":
    main()
