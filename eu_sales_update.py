#!/usr/bin/env python3

import os
import argparse
import calendar
import requests
import pandas as pd
import camelot
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(BASE_DIR, 'acea_pdfs')
DATA_DIR = os.path.join(BASE_DIR, 'data')

os.makedirs(PDF_DIR, exist_ok=True)

HEADERS = {"User-Agent": "Mozilla/5.0"}

def fetch_europe_pdf(year_month: str, url: str) -> str:
    year, mon = year_month.split('-')
    month_name = calendar.month_name[int(mon)]
    fname = f"{year}_{month_name}.pdf"
    path = os.path.join(PDF_DIR, fname)

    if os.path.exists(path):
        print(f"[Info] 이미 다운로드됨: {fname}")
        return path

    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    with open(path, 'wb') as f:
        f.write(resp.content)
    print(f"[Downloaded] {fname}")
    return path

def parse_europe_page6(pdf_path: str) -> pd.DataFrame:
    tables = camelot.read_pdf(
        pdf_path,
        pages='6',
        flavor='stream',
        strip_text='\n'
    )
    if not tables:
        raise RuntimeError(f"테이블 파싱 실패: {pdf_path}")
    df = tables[0].df
    df.columns = df.iloc[0]
    return df.iloc[1:].reset_index(drop=True)

def update_europe_sales_only(
    excel_file_arg: str,
    year_month: str,
    url: str,
    sheet_name: str = None
):
    if os.path.isabs(excel_file_arg):
        excel_path = excel_file_arg
    else:
        excel_path = os.path.join(DATA_DIR, excel_file_arg)

    pdf_file = fetch_europe_pdf(year_month, url)
    df6 = parse_europe_page6(pdf_file)
    target_sheet = sheet_name or year_month

    mode = 'a' if os.path.exists(excel_path) else 'w'
    writer_kwargs = {'engine': 'openpyxl', 'mode': mode}
    if mode == 'a':
        writer_kwargs['if_sheet_exists'] = 'new'

    with pd.ExcelWriter(excel_path, **writer_kwargs) as writer:
        df6.to_excel(writer, sheet_name=target_sheet, index=False)

    print(f"✅ '{excel_path}' 에 시트 '{target_sheet}' 추가 완료!")

def main():
    parser = argparse.ArgumentParser(
        prog='update_europe_sales',
        description='ACEA PDF(6페이지) → Excel 신규 시트로 추가'
    )
    parser.add_argument(
        'excel_file',
        help='data 폴더 내 파일명 또는 절대경로'
    )
    parser.add_argument(
        'year_month',
        help='추가할 연월 (YYYY-MM)'
    )
    parser.add_argument(
        'url',
        help='해당 연월의 ACEA PDF URL'
    )
    parser.add_argument(
        '-s', '--sheet',
        default=None,
        help='시트 이름 (기본: YYYY-MM)'
    )
    args = parser.parse_args()
    try:
        update_europe_sales_only(
            args.excel_file,
            args.year_month,
            args.url,
            sheet_name=args.sheet
        )
    except Exception as e:
        print(f"❌ 업데이트 실패: {e}")

if __name__ == '__main__':
    main()