#!/usr/bin/env python3
import argparse
import os
from datetime import datetime

import pandas as pd
import requests
from bs4 import BeautifulSoup

brand_map = {
    '比亚迪':       'BYD',
    '大众':         'Volkswagen',
    '吉利':         'Geely',
    '丰田':         'Toyota',
    '奇瑞':         'Chery',
    '长安':         'Changan',
    '本田':         'Honda',
    '特斯拉':       'Tesla',
    '奥迪':         'Audi',
    '五菱':         'Wuling',
    '五菱（银标）': 'Wuling (Silver)',
    '捷途':         'Jetour',
    '奔驰':         'Benz',
    '哈弗':         'Haval',
    'MG':           'MG',
    '宝马':         'BMW',
    '银河':         'GALAXY',
    '日产':         'Nissan',
    '红旗':         'Hongqi',
    '零跑':         'Leapmotor',
    '理想':         'LI',
    '别克':         'Buick',
    '小鹏':         'XPENG',
    '传祺':         'GAC Trumpchi',
    '小米':         'XIAOMI',
    '领克':         'Lynk & CO',
    '埃安':         'Aion',
    '起亚':         'Kia',
    '荣威':         'Roewe',
    '坦克':         'TANK',
    '现代':         'Hyundai',
}

def fetch_china_sales(ym: str, url: str) -> pd.DataFrame:
    """
    :param ym: 'YYYY-MM' 형식의 연월
    :param url: 해당 연월의 스크래핑 URL
    :return: index=Brand, columns=[Timestamp(연월 첫날)] 형태의 DataFrame
    """
    resp = requests.get(url, timeout=10)
    resp.encoding = resp.apparent_encoding
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "lxml")
    table = soup.find("table")
    if table is None:
        raise RuntimeError(f"🚨 {ym} 페이지에서 <table>을 찾을 수 없습니다: {url}")

    records = []
    for tr in table.find_all("tr")[1:]:
        tds = tr.find_all("td")
        if len(tds) < 3:
            continue
        cn = tds[1].get_text(strip=True)
        txt = tds[2].get_text(strip=True).replace(",", "")
        try:
            sales = int(txt)
        except ValueError:
            continue
        en = brand_map.get(cn)
        if en:
            records.append({"Brand": en, "Sales": sales})

    if not records:
        raise RuntimeError(f"🚨 {ym} 데이터가 없습니다: {url}")

    df = pd.DataFrame(records).set_index("Brand")
    month_ts = pd.to_datetime(f"{ym}-01")
    df.columns = [month_ts]
    return df

def normalize_columns_to_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    df.columns에 섞여 있는 str, datetime, Timestamp 형식을 모두 YYYY-MM-DD datetime 으로 변환
    """
    dates = pd.to_datetime(df.columns, errors='coerce')
    dates = dates.normalize()
    df.columns = dates
    return df

def update_china_sales_only(excel_path: str, ym: str, url: str, sheet_name: str = "Brands"):
    """
    :param excel_path: 업데이트할 Excel 파일 경로
    :param ym: 'YYYY-MM' 형식의 연월
    :param url: 해당 연월의 스크래핑 URL
    :param sheet_name: 대상 시트 이름 (기본 'Brands')
    """
    if os.path.exists(excel_path):
        df_existing = pd.read_excel(
            excel_path, sheet_name=sheet_name, index_col=0, engine="openpyxl"
        )
        df_existing = normalize_columns_to_date(df_existing)
    else:
        df_existing = pd.DataFrame()

    new_df = fetch_china_sales(ym, url)

    df_combined = pd.concat([df_existing, new_df], axis=1)
    df_combined = df_combined.loc[:, ~df_combined.columns.duplicated()]

    df_combined = df_combined.sort_index(axis=1)

    df_combined.to_excel(excel_path, sheet_name=sheet_name)
    print(f"✅ '{excel_path}' 파일에 중국 {ym} 데이터가 추가/업데이트 되었습니다.")

def main():
    parser = argparse.ArgumentParser(
        prog="chn_sales_update",
        description="Excel 파일에 중국 자동차 판매(월별) 데이터를 추가/업데이트 합니다."
    )
    parser.add_argument("excel_file", help="기존 엑셀 파일 경로")
    parser.add_argument("year_month", help="추가할 연월 (YYYY-MM)")
    parser.add_argument("url", help="해당 연월의 스크래핑 URL")
    parser.add_argument(
        "-s", "--sheet", default="Brands",
        help="대상 시트 이름 (기본: Brands)"
    )

    args = parser.parse_args()
    try:
        update_china_sales_only(
            args.excel_file, args.year_month, args.url, args.sheet
        )
    except Exception as e:
        print(f"❌ 업데이트 실패: {e}")

if __name__ == "__main__":
    main()