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

def insert_image_to_each_table(image_path, input_hwp_path, output_hwp_path):
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject.2")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

    hwp.Open(os.path.abspath(input_hwp_path))
    hwp.MovePos(3)  # 문서 처음으로 이동 


    head_ctrl = hwp.HeadCtrl
    found_table = False

    image_files = [
        f for f in os.listdir(image_path)
        if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))
    ]

    # 이미지 파일 목록 참조 (/output)
    for image_file in image_files:
            image_name = os.path.splitext(image_file)[0]
            if(image_name == "1번(취미)_1"):
                image_name = "규칙적인 여가 및 취미활동에 대한 결과"
            if(image_name == "2번(가사노동)_1"):
                image_name = "[하루 평균 가사노동시간에 대한 결과]"
            if(image_name == "3번(질병유무)_1"):
                image_name = "[질병 진단에 대한 결과]"
            if(image_name == "3-1번(질병유무, 질병종류)_1"):
                image_name = "[진단받은 질병명에 대한 결과]"
            if(image_name == "4번(사고)_1"):
                image_name = "[운동 중 혹은 사고로 신체 부위를 다친 적이 있는가에 대한 결과]"
            if(image_name == "4-1번(사고, 신체부위)_1"):
                image_name = "[운동 중 혹은 사고로 다친 신체 부위에 대한 결과]"
            if(image_name == "5번(장애)_1"):
                image_name = "[일의 육체적 부담 정도에 대한 결과]"
            if(image_name == "6번(근골증상여부)_1"):
                image_name = "[작업과 관련하여 통증이나 불편함을 느낀 적이 있는가에 대한 결과]"
            if(image_name == "6-1번(근골, 신체부위)_1"):
                image_name = "[통증의 구체적 부위에 대한 결과]"
            if(image_name == "6-2번(통증기간지속)_1"):
                image_name = "[통증의 지속 기간에 대한 결과]"
            if(image_name == "6-3번(통증정도)_1"):
                image_name = "[통증의 정도에 대한 결과]"
            if(image_name == "6-4번(통증빈도)_1"):
                image_name = "[통증의 빈도에 대한 결과]"
            if(image_name == "6-5번(지난1주일증상여부)_1"):
                image_name = "[지난 1주일 동안 통증의 여부에 대한 결과]"
            if(image_name == "6-6번(통증어떤일)_1"):
                image_name = "[지난 1년 동안 통증으로 인해 발생한 일에 대한 결과]"
            if(image_name == "6-7번(증상자분류)_1"):
                image_name = "[근골격계질환 요주의자/유소견자 추정에 대한 결과]"

            try:
                # 텍스트 검색
                hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                hwp.HParameterSet.HFindReplace.HSet.SetItem("FindString", image_name)
                hwp.HParameterSet.HFindReplace.HSet.SetItem("Direction", 1)
                hwp.HParameterSet.HFindReplace.HSet.SetItem("FindType", 1)
                found = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

                if not found:
                    print(f"❌ [{image_name}] 텍스트를 문서에서 찾지 못했습니다.")
                    continue

                ctrl = hwp.ParentCtrl
                if ctrl.CtrlID != "tbl":
                    print(f"⚠️ [{image_name}] 텍스트는 찾았지만 표 안이 아닙니다.")
                    continue

                # 표의 첫 번째 셀 선택
                hwp.MoveToField(image_name, True, True, False)
                # 현재 캐럿이 위치한 셀에서 열(column)의 시작
                hwp.MovePos(106)
                hwp.InsertPicture(os.path.join(image_path, image_file), True, 2)

                print(f"✅ [{image_name}] 표에 이미지 삽입 완료")
                found_table = True

            except Exception as e:
                print(f"⚠️ [{image_name}] 처리 중 오류 발생: {e}")

    if not found_table:
        print("🔍 문서 내 표를 찾지 못했거나 삽입할 수 없었습니다.")

    # 저장
    hwp.SaveAs(os.path.abspath(output_hwp_path))
    hwp.Quit()
    del hwp
    print(f"💾 작업 완료. 저장 경로: {output_hwp_path}")


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
