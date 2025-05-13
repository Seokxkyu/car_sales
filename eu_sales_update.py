#!/usr/bin/env python3
import os
import argparse
import calendar
import requests
import pandas as pd
import camelot
from openpyxl import load_workbook

SAVE_PDF_DIR = "acea_pdfs"
os.makedirs(SAVE_PDF_DIR, exist_ok=True)
HEADERS = {"User-Agent": "Mozilla/5.0"}

def fetch_europe_pdf(year_month: str, url: str) -> str:
    """
    :param year_month: 'YYYY-MM' 형식
    :param url: ACEA PDF URL
    :returns: 로컬에 저장된 PDF 경로
    """
    year, mon = year_month.split("-")
    month_name = calendar.month_name[int(mon)]
    fname = f"{year}_{month_name}.pdf"
    path = os.path.join(SAVE_PDF_DIR, fname)

    if os.path.exists(path):
        print(f"[Info] 이미 다운로드됨: {fname}")
        return path

    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    with open(path, "wb") as f:
        f.write(resp.content)
    print(f"[Downloaded] {fname}")
    return path

def parse_europe_page6(pdf_path: str) -> pd.DataFrame:
    """
    :param pdf_path: fetch_europe_pdf()로 받은 PDF 파일 경로
    :returns: 6페이지 첫 테이블을 DataFrame 으로 반환 (헤더+본문)
    """
    tables = camelot.read_pdf(
        pdf_path,
        pages="6",
        flavor="stream",
        strip_text="\n"
    )
    if not tables:
        raise RuntimeError(f"테이블 파싱 실패: {pdf_path}")
    df = tables[0].df
    df.columns = df.iloc[0]
    return df.iloc[1:].reset_index(drop=True)

def update_europe_sales_only(
    excel_path: str,
    year_month: str,
    url: str,
    sheet_name: str = None
):
    """
    :param excel_path: 업데이트할 Excel 파일 경로
    :param year_month: 'YYYY-MM' 형식
    :param url: 해당 연월의 PDF URL
    :param sheet_name: 시트 이름 (기본: 'YYYY-MM')
    """
    pdf_file = fetch_europe_pdf(year_month, url)

    df6 = parse_europe_page6(pdf_file)

    target_sheet = sheet_name or year_month

    mode = 'a' if os.path.exists(excel_path) else 'w'
    with pd.ExcelWriter(
        excel_path,
        engine="openpyxl",
        mode=mode,
        if_sheet_exists="new"
    ) as writer:
        df6.to_excel(writer, sheet_name=target_sheet, index=False)
    print(f"✅ '{excel_path}' 에 시트 '{target_sheet}' 추가 완료!")

def main():
    parser = argparse.ArgumentParser(
        prog="update_europe_sales",
        description="ACEA PDF(6페이지) → Excel 신규 시트로 추가"
    )
    parser.add_argument(
        "excel_file",
        help="대상 Excel 파일 경로 (존재하지 않으면 새로 생성)"
    )
    parser.add_argument(
        "year_month",
        help="추가할 연월 (YYYY-MM)"
    )
    parser.add_argument(
        "url",
        help="해당 연월의 ACEA PDF URL"
    )
    parser.add_argument(
        "-s", "--sheet",
        help="시트 이름 (기본: YYYY-MM)",
        default=None
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

if __name__ == "__main__":
    main()
