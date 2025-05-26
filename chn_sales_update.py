#!/usr/bin/env python3
import argparse
import os
import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

brand_map = {
    '比亚迪':'BYD','大众':'VW','吉利':'Geely','丰田':'Toyota','奇瑞':'Chery',
    '长安':'Changan','本田':'Honda','特斯拉':'Tesla','奥迪':'Audi','五菱':'Wuling',
    '五菱（银标）':'Wuling (Silver)','捷途':'Jetour','奔驰':'Benz','哈弗':'Haval',
    'MG':'MG','宝马':'BMW','银河':'GALAXY','日产':'Nissan','红旗':'Hongqi',
    '零跑':'Leapmotor','理想':'LI','别克':'Buick','小鹏':'XPENG','传祺':'GAC Trumpchi',
    '小米':'XIAOMI','领克':'Lynk&CO','埃安':'Aion','起亚':'Kia','荣威':'Roewe',
    '坦克':'TANK','现代':'Hyundai'
}

def fetch_china_sales(ym, url):
    resp = requests.get(url, timeout=10)
    resp.encoding = resp.apparent_encoding
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'lxml')
    table = soup.find('table')
    if not table:
        raise RuntimeError(f"{ym} 페이지에서 테이블을 찾을 수 없습니다")
    records = []
    for tr in table.find_all('tr')[1:]:
        tds = tr.find_all('td')
        if len(tds) < 3:
            continue
        cn = tds[1].get_text(strip=True)
        txt = tds[2].get_text(strip=True).replace(',', '')
        try:
            sales = int(txt)
        except ValueError:
            continue
        en = brand_map.get(cn)
        if en:
            records.append({'Brand': en, 'Sales': sales})
    if not records:
        raise RuntimeError(f"{ym} 데이터가 없습니다")
    df = pd.DataFrame(records).set_index('Brand')
    df.columns = [pd.to_datetime(f"{ym}-01").normalize()]
    return df

def normalize_columns_to_date(df):
    df.columns = pd.to_datetime(df.columns, errors='coerce').normalize()
    return df

def update_china_sales_only(excel_file_arg, ym, url, sheet_name='Brands'):
    if os.path.isabs(excel_file_arg):
        excel_path = excel_file_arg
    else:
        excel_path = os.path.join(DATA_DIR, excel_file_arg)
    if os.path.exists(excel_path):
        df_existing = pd.read_excel(excel_path, sheet_name=sheet_name, index_col=0, engine='openpyxl')
        df_existing = normalize_columns_to_date(df_existing)
    else:
        df_existing = pd.DataFrame()
    new_df = fetch_china_sales(ym, url)
    df_combined = pd.concat([df_existing, new_df], axis=1)
    df_combined = df_combined.loc[:, ~df_combined.columns.duplicated()]
    df_combined = df_combined.sort_index(axis=1)
    df_combined.to_excel(excel_path, sheet_name=sheet_name)
    print(f"'{excel_path}' 업데이트 완료")

def main():
    parser = argparse.ArgumentParser(prog='chn_sales_update', description='중국 자동차 판매 업데이트')
    parser.add_argument('excel_file', help='data 폴더 내 파일명 또는 절대경로')
    parser.add_argument('year_month', help='YYYY-MM')
    parser.add_argument('url', help='스크래핑 URL')
    parser.add_argument('-s', '--sheet', default='Brands', help='시트 이름')
    args = parser.parse_args()
    try:
        update_china_sales_only(args.excel_file, args.year_month, args.url, args.sheet)
    except Exception as e:
        print(f"업데이트 실패: {e}")

if __name__ == '__main__':
    main()