# #!/usr/bin/env python3

# import os
# import argparse
# import calendar
# import requests
# import pandas as pd
# import camelot
# from openpyxl import load_workbook

# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# PDF_DIR = os.path.join(BASE_DIR, 'acea_pdfs')
# DATA_DIR = os.path.join(BASE_DIR, 'data')

# os.makedirs(PDF_DIR, exist_ok=True)

# HEADERS = {"User-Agent": "Mozilla/5.0"}

# def fetch_europe_pdf(year_month: str, url: str) -> str:
#     year, mon = year_month.split('-')
#     month_name = calendar.month_name[int(mon)]
#     fname = f"{year}_{month_name}.pdf"
#     path = os.path.join(PDF_DIR, fname)

#     if os.path.exists(path):
#         print(f"[Info] 이미 다운로드됨: {fname}")
#         return path

#     resp = requests.get(url, headers=HEADERS, timeout=30)
#     resp.raise_for_status()
#     with open(path, 'wb') as f:
#         f.write(resp.content)
#     print(f"[Downloaded] {fname}")
#     return path

# def parse_europe_page6(pdf_path: str) -> pd.DataFrame:
#     tables = camelot.read_pdf(
#         pdf_path,
#         pages='6',
#         flavor='stream',
#         strip_text='\n'
#     )
#     if not tables:
#         raise RuntimeError(f"테이블 파싱 실패: {pdf_path}")
#     df = tables[0].df
#     df.columns = df.iloc[0]
#     return df.iloc[1:].reset_index(drop=True)

# def update_europe_sales_only(
#     excel_file_arg: str,
#     year_month: str,
#     url: str,
#     sheet_name: str = None
# ):
#     if os.path.isabs(excel_file_arg):
#         excel_path = excel_file_arg
#     else:
#         excel_path = os.path.join(DATA_DIR, excel_file_arg)

#     pdf_file = fetch_europe_pdf(year_month, url)
#     df6 = parse_europe_page6(pdf_file)
#     target_sheet = sheet_name or year_month

#     mode = 'a' if os.path.exists(excel_path) else 'w'
#     writer_kwargs = {'engine': 'openpyxl', 'mode': mode}
#     if mode == 'a':
#         writer_kwargs['if_sheet_exists'] = 'new'

#     with pd.ExcelWriter(excel_path, **writer_kwargs) as writer:
#         df6.to_excel(writer, sheet_name=target_sheet, index=False)

#     print(f"✅ '{excel_path}' 에 시트 '{target_sheet}' 추가 완료!")

# def main():
#     parser = argparse.ArgumentParser(
#         prog='update_europe_sales',
#         description='ACEA PDF(6페이지) → Excel 신규 시트로 추가'
#     )
#     parser.add_argument(
#         'excel_file',
#         help='data 폴더 내 파일명 또는 절대경로'
#     )
#     parser.add_argument(
#         'year_month',
#         help='추가할 연월 (YYYY-MM)'
#     )
#     parser.add_argument(
#         'url',
#         help='해당 연월의 ACEA PDF URL'
#     )
#     parser.add_argument(
#         '-s', '--sheet',
#         default=None,
#         help='시트 이름 (기본: YYYY-MM)'
#     )
#     args = parser.parse_args()
#     try:
#         update_europe_sales_only(
#             args.excel_file,
#             args.year_month,
#             args.url,
#             sheet_name=args.sheet
#         )
#     except Exception as e:
#         print(f"❌ 업데이트 실패: {e}")

# if __name__ == '__main__':
#     main()

#!/usr/bin/env python3

import os
import argparse
import calendar
import requests
import pandas as pd
import tabula
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
    # tabula-py를 이용해 6페이지 첫 테이블 읽기
    tables = tabula.read_pdf(
        pdf_path,
        pages='6',
        multiple_tables=True,
        stream=True
    )
    if not tables:
        raise RuntimeError(f"테이블 파싱 실패: {pdf_path}")
    df = tables[0]
    df.columns = df.iloc[0]
    return df.iloc[1:].reset_index(drop=True)


def update_europe_sales_only(
    excel_file_arg: str,
    year_month: str,
    url: str,
    sheet_name: str = None
):
    # 엑셀 경로 결정
    if os.path.isabs(excel_file_arg):
        excel_path = excel_file_arg
    else:
        excel_path = os.path.join(DATA_DIR, excel_file_arg)

    # PDF 다운로드 및 파싱
    pdf_file = fetch_europe_pdf(year_month, url)
    df6 = parse_europe_page6(pdf_file)

    # 시트명 및 mode 설정
    target_sheet = sheet_name or year_month
    mode = 'a' if os.path.exists(excel_path) else 'w'
    writer_kwargs = {'engine': 'openpyxl', 'mode': mode}
    if mode == 'a':
        writer_kwargs['if_sheet_exists'] = 'new'

    # 전처리: 세 번째 컬럼(split 대상)의 첫 번째 값만 추출
    col_c = df6.columns[2]
    new_col = target_sheet
    df6[new_col] = df6[col_c].astype(str).str.split(' ').str[0]
    df6.drop(columns=[col_c], inplace=True)

    # 엑셀에 시트 추가 (헤더 없이 데이터만)
    with pd.ExcelWriter(excel_path, **writer_kwargs) as writer:
        df6.to_excel(writer, sheet_name=target_sheet, index=False, header=False)

    print(f"✅ '{excel_path}' 에 시트 '{target_sheet}' 추가 완료!")


def main():
    parser = argparse.ArgumentParser(
        prog='update_europe_sales',
        description='ACEA PDF(6페이지) → Excel 신규 시트로 추가 (tabula 사용)'
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