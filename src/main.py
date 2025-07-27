import tkinter as tk 
from tkinter import filedialog
import win32com.client
import os
from utils.get_absolute_path import get_absolute_path

def choose_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()  # GUI 창 숨김
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return file_path

def choose_hwp_file():
    hwp_path = choose_file("HWP 파일 선택", [("HWP files", "*.hwp"), ("HWPX files", "*.hwpx")])
    if not hwp_path:
            print("❌ HWP 파일이 선택되지 않았습니다.")
            return

    try:
        # 4. 엑셀에서 차트 추출
        print(f"\n🔄 4단계: 엑셀 차트 추출 중...")
        chart_count = export_excel_charts(xlsx_path, output_dir)
        print(f"✅ {chart_count}개의 차트를 추출했습니다.")

        # 5. 한글 파일 열기
        print(f"\n🔄 5단계: 한글 파일 열기...")
        hwp = read_hwp_file(hwp_path)
        print(f"✅ HWP 파일을 성공적으로 열었습니다: {hwp_path}")

        # 6. 이미지-도형 매칭 및 삽입
        ## print(f"\n🔄 6단계: 이미지-도형 매칭 및 삽입...")
        ## match_images_to_shapes(hwp, os.path.join(os.path.dirname(hwp_path), "output"))

        # 6-1. 이미지 저장 경로에 있는 이미지 모두 한글에 삽입
        print(f"\n🔄 6-1단계: 이미지 한글에 삽입...")
        for img in os.listdir(output_dir):
            if img.endswith(".png"):
                insert_image_to_hwp(hwp_path, os.path.join(output_dir, img))

        # 7. 파일 저장 및 종료
        print(f"\n🔄 7단계: 파일 저장...")
        hwp.Save()
        hwp.Quit()
        print(f"✅ 작업이 완료되었습니다!")

    except Exception as e:
        print(f"❌ 작업 중 오류 발생: {e}")
    return choose_hwp_file()

def main():
    print("=" * 60)
    print("📊 엑셀 차트 → 한글 도형 삽입 프로그램")
    print("=" * 60)

    # 1. 엑셀 파일 선택
    print("\n📈 1단계: 엑셀 파일 선택")
    xlsx_path = choose_file("엑셀 파일 선택", [("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
    if not xlsx_path:
        print("❌ 엑셀 파일이 선택되지 않았습니다.")
        return

    # 2. 이미지 저장 폴더 선택(기본 값으로 대체, 유저 선택 없음)
    temp_image_path = get_absolute_path('./output')

    # 3. 한글 파일 선택
    choose_hwp_file()

if __name__ == "__main__":
    main()
