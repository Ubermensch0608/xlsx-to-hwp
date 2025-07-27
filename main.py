import tkinter as tk 
from tkinter import filedialog
import win32com.client
import os

def choose_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()  # GUI 창 숨김
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return file_path


def choose_folder(title="이미지를 저장할 폴더를 선택하세요"):
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    return folder_path

def insert_image_to_shape(hwp, image_path, shape_caption):
    """특정 캡션을 가진 도형 안에 이미지를 삽입합니다."""
    try:
        # 문서 시작으로 이동
        hwp.MovePos(2)
        
        # 페이지별로 도형 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 도형 검색
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 도형 찾기 시도 (다양한 방법)
                        found = False
                        
                        # 방법 1: FindCtrl 사용
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("ShapeObject")
                            except:
                                pass
                        
                        # 방법 2: FindCtrlID 사용
                        if not found and hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("ShapeObject")
                            except:
                                pass
                        
                        if found:
                            try:
                                # 도형의 캡션 정보 가져오기
                                if hasattr(hwp, 'GetCtrlData'):
                                    shape = hwp.GetCtrlData("ShapeObject")
                                    
                                    if hasattr(shape, 'Caption') and shape.Caption == shape_caption:
                                        print(f"✅ 캡션 '{shape_caption}'을 가진 도형을 찾았습니다.")
                                        
                                        # 도형 내부로 이동
                                        if hasattr(hwp, 'MoveToCtrl'):
                                            hwp.MoveToCtrl("ShapeObject")
                                        
                                        # 이미지 삽입
                                        if hasattr(hwp, 'InsertPicture'):
                                            hwp.InsertPicture(image_path, Embedded=True)
                                            print(f"✅ 이미지 삽입 완료: {image_path}")
                                            return True
                            except Exception as e:
                                print(f"⚠️ 도형 처리 중 오류: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 도형 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
        
        print(f"❌ 캡션 '{shape_caption}'을 가진 도형을 찾을 수 없습니다.")
        return False
        
    except Exception as e:
        print(f"❌ 이미지 삽입 중 오류: {e}")
        return False

def get_shape_captions(hwp):
    """HWP 파일의 모든 도형 캡션을 가져옵니다."""
    captions = []
    
    try:
        # 페이지별로 도형 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 도형 검색
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 도형 찾기 시도 (다양한 방법)
                        found = False

                        # 방법 1: FindCtrl 사용
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("Shape")
                            except:
                                pass
                        
                        # 방법 2: FindCtrlID 사용
                        if not found and hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("Shape")
                            except:
                                pass
                        
                        if found:
                            try:
                                # 도형의 캡션 정보 가져오기
                                if hasattr(hwp, 'GetCtrlData'):
                                    shape = hwp.GetCtrlData("Shape")

                                    if hasattr(shape, 'Caption') and shape.Caption:
                                        captions.append(shape.Caption)
                                        print(f"📋 페이지 {page_num + 1} - 발견된 도형 캡션: {shape.Caption}")
                            except Exception as e:
                                print(f"⚠️ 도형 캡션 추출 중 오류: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 도형 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
        
        return captions
        
    except Exception as e:
        print(f"❌ 도형 캡션 추출 중 오류: {e}")
        return []

def match_images_to_shapes(hwp, image_dir):
    """이미지 파일과 도형 캡션을 매칭하여 삽입합니다."""
    print("=" * 60)
    print("🖼️ 이미지-도형 매칭 및 삽입")
    print("=" * 60)
    
    # 1. 도형 캡션 목록 가져오기
    print("📋 도형 캡션 추출 중...")
    shape_captions = get_shape_captions(hwp)
    
    if not shape_captions:
        print("❌ 도형 캡션을 찾을 수 없습니다.")
        return
    
    print(f"✅ 총 {len(shape_captions)}개의 도형 캡션을 찾았습니다.")
    
    # 2. 이미지 파일 목록 가져오기
    print("\n🖼️ 이미지 파일 검색 중...")
    image_files = []
    
    try:
        for file in os.listdir(image_dir):
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                image_files.append(file)
                print(f"📸 발견된 이미지: {file}")
    except Exception as e:
        print(f"❌ 이미지 디렉토리 접근 오류: {e}")
        return
    
    if not image_files:
        print("❌ 이미지 파일을 찾을 수 없습니다.")
        return
    
    print(f"✅ 총 {len(image_files)}개의 이미지 파일을 찾았습니다.")
    
    # 3. 이미지와 도형 캡션 매칭
    print("\n🔗 이미지-도형 매칭 중...")
    matched_count = 0
    
    for image_file in image_files:
        # 이미지 파일명에서 확장자 제거
        image_name = os.path.splitext(image_file)[0]
        
        # 도형 캡션과 매칭 시도
        for caption in shape_captions:
            if image_name.lower() in caption.lower() or caption.lower() in image_name.lower():
                print(f"\n🎯 매칭 발견: '{image_name}' ↔ '{caption}'")
                
                # 이미지 삽입
                image_path = os.path.join(image_dir, image_file)
                if insert_image_to_shape(hwp, image_path, caption):
                    matched_count += 1
                    break
        else:
            print(f"⚠️ '{image_name}'에 매칭되는 도형 캡션을 찾을 수 없습니다.")
    
    print(f"\n✅ 총 {matched_count}개의 이미지가 성공적으로 삽입되었습니다.")

def export_excel_charts(xlsx_path, output_dir):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(xlsx_path)

    chart_count = 0
    for i, sheet in enumerate(wb.Sheets):
        for j, chart in enumerate(sheet.ChartObjects()):
            chart_obj = chart.Chart
            image_path = os.path.join(output_dir, f"차트{i+1}-{j+1}.png")
            chart_obj.Export(image_path)
            print(f"[✔] 차트 이미지 저장됨: {image_path}")
            chart_count += 1
    
    wb.Close(False)
    excel.Quit()
    return chart_count

def insert_image_to_hwp(hwp_path, image_path, position=0):
    hwp = read_hwp_file(hwp_path)
 
    hwp.MovePos(position)  # 문서 시작: 2, 끝: 3 등
    hwp.InsertPicture(image_path, Embedded=True)

    print(f"[✔] 이미지 삽입됨: {image_path}")

def read_hwp_file(hwp_path):
    hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안 DLL 등록
    hwp.Open(hwp_path)
    hwp.XHwpWindows.Item(0).Visible = True
    return hwp

def get_hwp_object_info(hwp):
    """HWP 파일의 내부 객체 정보를 추출합니다."""
    print("=" * 50)
    print("📄 HWP 파일 객체 정보 분석")
    print("=" * 50)
    
    # 문서 전체 정보
    print(f"📋 문서 정보:")
    try:
        print(f"  - 페이지 수: {hwp.PageCount}")
        
        # 총 글자 수 계산
        total_chars = 0
        for i in range(hwp.PageCount):
            try:
                # 페이지 이동 시도
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(i)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(i)
                else:
                    hwp.MovePos(2)  # 문서 시작으로 이동
                
                page_text = hwp.GetPageText(i)
                total_chars += len(page_text)
            except:
                pass
        print(f"  - 총 글자 수: {total_chars}")
        
        # 총 줄 수 계산
        total_lines = 0
        for i in range(hwp.PageCount):
            try:
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(i)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(i)
                else:
                    hwp.MovePos(2)
                
                page_text = hwp.GetPageText(i)
                total_lines += page_text.count('\n') + 1
            except:
                pass
        print(f"  - 총 줄 수: {total_lines}")
        
    except Exception as e:
        print(f"  - 문서 정보 추출 중 오류: {e}")
    
    # 도형 객체 정보
    print(f"\n🔷 도형 객체 정보:")
    try:
        shape_count = 0
        
        # 페이지별로 도형 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 도형 검색
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 도형 찾기 시도 (다양한 방법)
                        found = False
                        
                        # 방법 1: FindCtrl 사용
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("ShapeObject")
                            except:
                                pass
                        
                        # 방법 2: FindCtrlID 사용
                        if not found and hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("ShapeObject")
                            except:
                                pass
                        
                        if found:
                            shape_count += 1
                            print(f"  - 페이지 {page_num + 1}의 도형 {shape_count}")
                            
                            # 도형의 속성 정보
                            try:
                                if hasattr(hwp, 'GetCtrlData'):
                                    shape = hwp.GetCtrlData("ShapeObject")
                                    if hasattr(shape, 'Caption'):
                                        print(f"    캡션: {shape.Caption}")
                                    if hasattr(shape, 'Name'):
                                        print(f"    이름: {shape.Name}")
                                    if hasattr(shape, 'Type'):
                                        print(f"    타입: {shape.Type}")
                            except Exception as e:
                                print(f"    도형 속성 정보를 가져올 수 없습니다: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 도형 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
        
        if shape_count == 0:
            print("  - 도형 객체가 없습니다.")
            
    except Exception as e:
        print(f"  - 도형 정보 추출 중 오류: {e}")
    
    # 이미지 객체 정보
    print(f"\n🖼️ 이미지 객체 정보:")
    try:
        image_count = 0
        
        # 페이지별로 이미지 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 이미지 검색
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 이미지 찾기 시도
                        found = False
                        
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("Picture")
                            except:
                                pass
                        
                        if hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("Picture")
                            except:
                                pass
                        
                        if found:
                            image_count += 1
                            print(f"  - 페이지 {page_num + 1}의 이미지 {image_count}")
                            
                            # 이미지의 속성 정보
                            try:
                                if hasattr(hwp, 'GetCtrlData'):
                                    pic = hwp.GetCtrlData("Picture")
                                    if hasattr(pic, 'Name'):
                                        print(f"    이름: {pic.Name}")
                                    if hasattr(pic, 'Width'):
                                        print(f"    너비: {pic.Width}")
                                    if hasattr(pic, 'Height'):
                                        print(f"    높이: {pic.Height}")
                            except Exception as e:
                                print(f"    이미지 속성 정보를 가져올 수 없습니다: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 이미지 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
        
        if image_count == 0:
            print("  - 이미지 객체가 없습니다.")
            
    except Exception as e:
        print(f"  - 이미지 정보 추출 중 오류: {e}")
    
    # 표 객체 정보
    print(f"\n📊 표 객체 정보:")
    try:
        table_count = 0
        
        # 페이지별로 표 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 표 검색
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 표 찾기 시도
                        found = False
                        
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("Table")
                            except:
                                pass
                        
                        if hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("Table")
                            except:
                                pass
                        
                        if found:
                            table_count += 1
                            print(f"  - 페이지 {page_num + 1}의 표 {table_count}")
                            
                            # 표의 속성 정보
                            try:
                                if hasattr(hwp, 'GetCtrlData'):
                                    table = hwp.GetCtrlData("Table")
                                    if hasattr(table, 'RowCount'):
                                        print(f"    행 수: {table.RowCount}")
                                    if hasattr(table, 'ColCount'):
                                        print(f"    열 수: {table.ColCount}")
                            except Exception as e:
                                print(f"    표 속성 정보를 가져올 수 없습니다: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 표 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
        
        if table_count == 0:
            print("  - 표 객체가 없습니다.")
            
    except Exception as e:
        print(f"  - 표 정보 추출 중 오류: {e}")
    
    # 텍스트 상자 객체 정보
    print(f"\n📝 텍스트 상자 객체 정보:")
    try:
        textbox_count = 0
        
        # 페이지별로 텍스트 상자 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 텍스트 상자 검색
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 텍스트 상자 찾기 시도
                        found = False
                        
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("TextBox")
                            except:
                                pass
                        
                        if hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("TextBox")
                            except:
                                pass
                        
                        if found:
                            textbox_count += 1
                            print(f"  - 페이지 {page_num + 1}의 텍스트 상자 {textbox_count}")
                            
                            # 텍스트 상자의 속성 정보
                            try:
                                if hasattr(hwp, 'GetCtrlData'):
                                    textbox = hwp.GetCtrlData("TextBox")
                                    if hasattr(textbox, 'Text'):
                                        print(f"    내용: {textbox.Text[:50]}...")  # 처음 50자만 표시
                            except Exception as e:
                                print(f"    텍스트 상자 속성 정보를 가져올 수 없습니다: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 텍스트 상자 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
        
        if textbox_count == 0:
            print("  - 텍스트 상자 객체가 없습니다.")
            
    except Exception as e:
        print(f"  - 텍스트 상자 정보 추출 중 오류: {e}")

def get_detailed_shape_info(hwp):
    """도형의 상세 정보를 추출합니다."""
    print(f"\n🔍 도형 상세 정보:")
    print("-" * 30)
    
    try:
        # 페이지별로 도형 검색
        for page_num in range(hwp.PageCount):
            try:
                # 페이지로 이동
                if hasattr(hwp, 'MoveToPageIndex'):
                    hwp.MoveToPageIndex(page_num)
                elif hasattr(hwp, 'MoveToPage'):
                    hwp.MoveToPage(page_num)
                
                # 페이지 내에서 도형 검색
                shape_index = 0
                max_iterations = 100
                iteration = 0
                
                while iteration < max_iterations:
                    try:
                        # 도형 찾기 시도 (다양한 방법)
                        found = False
                        
                        # 방법 1: FindCtrl 사용
                        if hasattr(hwp, 'FindCtrl'):
                            try:
                                found = hwp.FindCtrl("ShapeObject")
                            except:
                                pass
                        
                        # 방법 2: FindCtrlID 사용
                        if not found and hasattr(hwp, 'FindCtrlID'):
                            try:
                                found = hwp.FindCtrlID("ShapeObject")
                            except:
                                pass
                        
                        if found:
                            shape_index += 1
                            print(f"\n📐 페이지 {page_num + 1} - 도형 {shape_index}")
                            
                            # 도형의 기본 정보
                            try:
                                if hasattr(hwp, 'GetCtrlData'):
                                    shape = hwp.GetCtrlData("ShapeObject")
                                    
                                    # 캡션 정보 추출
                                    if hasattr(shape, 'Caption'):
                                        caption = shape.Caption
                                        print(f"  캡션: {caption}")
                                    
                                    # 도형 타입 정보
                                    if hasattr(shape, 'ShapeType'):
                                        shape_type = shape.ShapeType
                                        print(f"  도형 타입: {shape_type}")
                                    
                                    # 도형 크기 정보
                                    if hasattr(shape, 'Width') and hasattr(shape, 'Height'):
                                        print(f"  크기: {shape.Width} x {shape.Height}")
                                    
                                    # 도형 위치 정보
                                    if hasattr(shape, 'Left') and hasattr(shape, 'Top'):
                                        print(f"  위치: ({shape.Left}, {shape.Top})")
                                    
                                    # 도형 이름
                                    if hasattr(shape, 'Name'):
                                        print(f"  이름: {shape.Name}")
                                    
                                    # 도형 내부 텍스트 (있는 경우)
                                    try:
                                        if hasattr(hwp, 'MoveToCtrl'):
                                            hwp.MoveToCtrl("ShapeObject")
                                        if hasattr(hwp, 'GetText'):
                                            shape_text = hwp.GetText()
                                            if shape_text.strip():
                                                print(f"  내부 텍스트: {shape_text[:100]}...")
                                    except:
                                        pass
                            except Exception as e:
                                print(f"  상세 정보 추출 실패: {e}")
                        
                        # 다음 컨트롤로 이동
                        if hasattr(hwp, 'MoveNextCtrl'):
                            hwp.MoveNextCtrl()
                        else:
                            break
                            
                        iteration += 1
                        
                    except Exception as e:
                        print(f"⚠️ 페이지 {page_num + 1} 도형 검색 중 오류: {e}")
                        break
                
            except Exception as e:
                print(f"⚠️ 페이지 {page_num + 1} 이동 중 오류: {e}")
                continue
                
    except Exception as e:
        print(f"도형 상세 정보 추출 중 오류: {e}")

def explore_hwp_object(hwp):
    """HWP 객체의 사용 가능한 메서드와 프로퍼티를 탐색합니다."""
    print("=" * 60)
    print("🔍 HWP 객체 메서드 및 프로퍼티 탐색")
    print("=" * 60)
    
    # 객체 타입 정보
    print(f"📋 객체 타입: {type(hwp)}")
    print(f"📋 객체 문자열 표현: {str(hwp)}")
    
    # 모든 속성과 메서드 목록
    print(f"\n🔧 사용 가능한 속성 및 메서드:")
    print("-" * 40)
    
    # dir() 함수로 모든 속성과 메서드 가져오기
    all_attributes = dir(hwp)
    
    # 내장 속성/메서드 제외하고 HWP 관련 것만 필터링
    hwp_attributes = []
    for attr in all_attributes:
        if not attr.startswith('_') and not attr.startswith('__'):
            hwp_attributes.append(attr)
    
    # 알파벳 순으로 정렬
    hwp_attributes.sort()
    
    # 속성과 메서드 분류
    properties = []
    methods = []
    
    for attr in hwp_attributes:
        try:
            attr_obj = getattr(hwp, attr)
            if callable(attr_obj):
                methods.append(attr)
            else:
                properties.append(attr)
        except:
            # 접근할 수 없는 속성은 메서드로 간주
            methods.append(attr)
    
    # 프로퍼티 출력
    print(f"\n📊 프로퍼티 ({len(properties)}개):")
    for i, prop in enumerate(properties, 1):
        try:
            value = getattr(hwp, prop)
            print(f"  {i:2d}. {prop} = {value}")
        except Exception as e:
            print(f"  {i:2d}. {prop} = [접근 오류: {e}]")
    
    # 메서드 출력
    print(f"\n⚙️ 메서드 ({len(methods)}개):")
    for i, method in enumerate(methods, 1):
        print(f"  {i:2d}. {method}()")
    
    # 주요 메서드 상세 정보
    print(f"\n🔍 주요 메서드 상세 정보:")
    print("-" * 40)
    
    important_methods = [
        'GetPageText', 'MovePos', 'FindCtrl', 'GetCtrlData', 
        'MoveToPage', 'MoveToPageIndex', 'MoveToPageTop',
        'SelectCtrl', 'MoveNextCtrl', 'MoveToCtrl', 'GetText'
    ]
    
    for method_name in important_methods:
        if hasattr(hwp, method_name):
            method = getattr(hwp, method_name)
            print(f"✅ {method_name}() - 사용 가능")
            try:
                # 메서드의 도움말 정보 가져오기
                if hasattr(method, '__doc__') and method.__doc__:
                    doc = method.__doc__.strip()
                    if doc:
                        print(f"   📝 설명: {doc[:100]}...")
            except:
                pass
        else:
            print(f"❌ {method_name}() - 사용 불가")
    
    # 객체의 특별한 속성들 확인
    print(f"\n🎯 특별한 속성 확인:")
    print("-" * 40)
    
    special_props = [
        'PageCount', 'CharCount', 'LineCount', 'Selection',
        'ActiveDocument', 'Documents', 'Application'
    ]
    
    for prop in special_props:
        if hasattr(hwp, prop):
            try:
                value = getattr(hwp, prop)
                print(f"✅ {prop} = {value}")
            except Exception as e:
                print(f"❌ {prop} = [접근 오류: {e}]")
        else:
            print(f"❌ {prop} - 속성이 없음")

def kill_process():
    print("▶️ ctrl + c 누르면 프롬그램 종료")
    
    try:
        while True:
            pass
    except KeyboardInterrupt:
        print("▶️ 프로그램 종료")
        exit()

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

    # 2. 이미지 저장 폴더 선택
    print("\n📁 2단계: 이미지 저장 폴더 선택")
    output_dir = choose_folder("이미지를 저장할 폴더를 선택하세요")
    if not output_dir:
        print("❌ 이미지 저장 폴더가 선택되지 않았습니다.")
        return

    # 3. 한글 파일 선택
    print("\n📄 3단계: 한글(HWP) 파일 선택")
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

if __name__ == "__main__":
    main()
