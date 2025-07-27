import sys
import win32com.client as win32
from utils.get_absolute_path import get_absolute_path
from utils.choose_file import choose_file
from utils.run_with_interrupt import run_with_interrupt

def export_excel_data():
    print("\n📈 1단계: 엑셀 파일 선택")
    
    xlsx_path = choose_file("엑셀 파일 선택", [("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
    if not xlsx_path:
        print("❌ 엑셀 파일이 선택되지 않았습니다.")
        sys.exit(1)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = excel.Workbooks.Open(xlsx_path)
    sheet = workbook.Sheets(1)

    try:
        for sheet_index, sheet in enumerate(workbook.Sheets):
            chart_count = sheet.ChartObjects().Count
            if chart_count == 0:
                continue

            for i in range(1, chart_count + 1):
                chart_obj = sheet.ChartObjects(i)
                chart = chart_obj.Chart

                filename = f"{sheet.Name}_{i}.png"
                filepath = get_absolute_path(f'./output/{filename}')

                chart.Export(Filename=filepath, FilterName="PNG")
                print(f"✅ 저장됨: {filepath}")

    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()
        del excel

def choose_hwp_file():
    hwp_path = choose_file("HWP 파일 선택", [("HWP files", "*.hwp"), ("HWPX files", "*.hwpx")])
    if not hwp_path:
        print("❌ HWP 파일이 선택되지 않았습니다.")
        sys.exit(1) 

def main():
    print("=" * 60)
    print("📊 엑셀 차트 → 한글 도형 삽입 프로그램")
    print("=" * 60)

    # 1. 엑셀 파일에서 데이터 추출
    # - 차트 데이터 추출
    # - 이미지 추출
    export_excel_data()
    
    # 2. 한글 파일 선택
    # choose_hwp_file()

    sys.exit(1)

if __name__ == "__main__":
    run_with_interrupt(main)
