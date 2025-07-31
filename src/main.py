import sys
import os
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

def insert_image_to_table_cell(hwp, table_name, row, col, image_path):
    hwp.MovePos(3)
    ctrl_count = hwp.GetCtrlCount()
    print(f"📦 컨트롤 개수: {ctrl_count}")

     # FindCtrl 초기화
    table_index = 0
    while True:
        try:
            hwp.HAction.GetDefault("FindCtrl", hwp.HParameterSet.HFindCtrl)
            found = hwp.HAction.Execute("FindCtrl", hwp.HParameterSet.HFindCtrl)

            if not found:
                break

            ctrl_name = hwp.HParameterSet.HFindCtrl.HCName
            ctrl_type = hwp.HParameterSet.HFindCtrl.HCType

            if ctrl_type != "tbl":
                continue  # 표가 아니면 스킵

            table_index += 1
            print(f"📋 [{table_index}] 표 발견: {ctrl_name}")

            # 표 셀 선택
            hwp.HAction.GetDefault("TableSelCell", hwp.HParameterSet.HTableSelCell)
            hwp.HParameterSet.HTableSelCell.TblIdx = 1  # 현재 선택된 표 기준
            hwp.HParameterSet.HTableSelCell.Row = 2
            hwp.HParameterSet.HTableSelCell.Col = 1
            hwp.HAction.Execute("TableSelCell", hwp.HParameterSet.HTableSelCell)

            # 이미지 삽입
            hwp.HAction.GetDefault("InsertPicture", hwp.HParameterSet.HInsertPicture.HSet)
            hwp.HParameterSet.HInsertPicture.filename = os.path.abspath(image_path)
            hwp.HParameterSet.HInsertPicture.KeepAspectRatio = 1
            hwp.HAction.Execute("InsertPicture", hwp.HParameterSet.HInsertPicture.HSet)
            print(f"✅ 이미지 삽입 완료 (표 {table_index})")
        
        except Exception as e:
            print(f"❌ 이미지 삽입 실패: {e}")
            sys.exit(1)


def insert_image_to_each_table(image_path, input_hwp_path, output_hwp_path):
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject.2")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

    hwp.Open(os.path.abspath(input_hwp_path))
    hwp.MovePos(3)  # 문서 처음으로 이동 

    # 저장
    hwp.SaveAs(os.path.abspath(output_hwp_path))
    hwp.Quit()
    print(f"🎉 완료! 저장된 파일: {output_hwp_path}")


def insert_data_to_hwp():
    print("\n📊 2단계: 한글 파일 선택")

    try:
        hwp_path = choose_file("HWP 파일 선택", [("HWP files", "*.hwp"), ("HWPX files", "*.hwpx")])
        file_dir = os.path.dirname(hwp_path)
        full_file_name = os.path.basename(hwp_path)
        [file_name, file_ext] = full_file_name.split('.')
        save_path = get_absolute_path(f'{file_dir}/{file_name}_result.{file_ext}')

        if not hwp_path:
            print("❌ HWP 파일이 선택되지 않았습니다.")
            sys.exit(1) 

        insert_image_to_each_table(get_absolute_path('./output'), hwp_path, save_path)

    except Exception as e:
        print(f"❌ 한글 파일 삽입 실패: {e}")
        sys.exit(1)
   
def main():
    print("=" * 60)
    print("📊 엑셀 차트 → 한글 도형 삽입 프로그램")
    print("=" * 60)

    # 1. 엑셀 파일에서 데이터 추출
    # - 차트 데이터 추출
    # - 이미지 추출
    # export_excel_data()
    
    # 2. 한글 파일에 데이터 삽입
    insert_data_to_hwp()

    print("\n✅ 작업이 완료되었습니다.")
    print("=" * 60)

if __name__ == "__main__":
    main()
