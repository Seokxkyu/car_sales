# 월간 차 판매량 업데이트

## 📂 디렉토리 구조
```
intern/
└── car_sales/
    ├── us_sales_update.py
    ├── chn_sales_update.py
    ├── eu_sales_update.py
    └── data/
        ├── us_sales_update.xlsx
        ├── china_sales_update.xlsx
        └── europe_sales_update.xlsx
```

## 사용법
### 1. Windows Powershell 관리자 권한으로 실행
### 2. intern/car_sales 폴더 진입
```bash
cd intern/car_sales
```
### 3. 미국 업데이트
**실행 방법**
```bash
python us_sales_update.py <OUTPUT_XLSX_PATH> --sheet <SHEET_NAME>
```
- `OUTPUT_XLSX_PATH`: 업데이트 하고자 하는 엑셀 파일 
- `SHEET_NAME`: Sheet 이름
- **예시** (매 월 동일)

    ```bash
    python us_sales_update.py "us_sales_update.xlsx" --sheet Brands
    ```


### 4. 중국 업데이트
**실행 방법**
```bash
python chn_sales_update.py <OUTPUT_XLSX_PATH> <YYYY-MM> <GASGOO_URL> --sheet <SHEET_NAME>
```
- `OUTPUT_XLSX_PATH`: 업데이트 하고자 하는 엑셀 파일 
- `SHEET_NAME`: Sheet 이름
- `<YYYY-MM>`: 연도-월
- `<GASGOO_URL>`: [Gasgoo 월별 중국 전 자동차 브랜드 판매 순위](https://auto.gasgoo.com/qcxl/article/76543.html)

- **예시** (2025년 3월 업데이트)
    ```bash
    python chn_sales_update.py "china_sales_update.xlsx" 2025-03 https://auto.gasgoo.com/qcxl/article/76543.html --sheet China
    ```


### 5. 유럽 업데이트
`Camelot` 라이브러리 사용을 위해 Python 3.11 환경 진입 필요
```bash
# 가상환경 설정
C:\Users\rootn\AppData\Local\Programs\Python\Python311\python.exe -m venv myenv311

# 가상환경 진입
.\myenv311\Scripts\activate.ps1
```

**실행 방법**
```bash
python eu_sales_update.py <PDF_URL> <YYYY-MM>
```
- `<PDF_URL>`: [ACEA 유럽 전 자동차 브랜드 판매 순위 포스트](https://www.acea.auto/nav/?content=press-releases)
- `<YYYY-MM>`: 연도-월

- **예시** (2025년 5월)
    ```bash
    python eu_sales_update.py https://www.acea.auto/files/Press_release_car_registrations_May_2025.pdf 2025-05
    ```