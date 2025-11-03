from pywinauto import application
import sys

# UTF-8 출력 설정
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

try:
    # win32 백엔드로 연결
    app = application.Application(backend="win32")
    app.connect(title="연말정산추가자료입력")
    dlg = app.window(title="연말정산추가자료입력")

    print("=== 탭 영역 내 모든 텍스트 검색 ===\n")

    # 탭 컨트롤 찾기
    tab_control = dlg.child_window(class_name="Afx:TabWnd:330000:8:10003:10", found_index=0)

    if tab_control.exists():
        print(f"탭 컨트롤 위치: {tab_control.rectangle()}\n")

        # 탭 컨트롤의 모든 자식 요소 검색
        all_children = tab_control.descendants()

        print(f"총 {len(all_children)}개의 자식 요소 발견\n")

        for i, child in enumerate(all_children[:30]):  # 처음 30개만
            try:
                text = child.window_text()
                class_name = child.class_name()
                rect = child.rectangle()

                if text:  # 텍스트가 있는 것만 출력
                    print(f"[{i}] 텍스트: '{text}'")
                    print(f"    클래스: {class_name}")
                    print(f"    위치: {rect}")
                    print()

                    # 부양가족 찾으면 클릭 시도
                    if "부양가족" in text:
                        print(f">>> '부양가족' 찾음! 클릭 시도...")
                        try:
                            child.click_input()
                            print("클릭 완료!")
                        except Exception as e:
                            print(f"클릭 실패: {e}")
                            # 대체 방법: 좌표로 클릭
                            from pywinauto import mouse
                            center_x = (rect.left + rect.right) // 2
                            center_y = (rect.top + rect.bottom) // 2
                            mouse.click(coords=(center_x, center_y))
                            print(f"좌표 ({center_x}, {center_y})로 클릭 완료!")
                        break

            except Exception as e:
                pass

    # 대체 방법: dlg의 모든 텍스트 검색
    print("\n=== 전체 윈도우에서 '부양가족' 텍스트 검색 ===\n")

    all_controls = dlg.descendants()
    found = False

    for ctrl in all_controls:
        try:
            text = ctrl.window_text()
            if text and "부양가족" in text and "불러오기" not in text:
                print(f"발견: '{text}' (클래스: {ctrl.class_name()})")
                rect = ctrl.rectangle()
                print(f"위치: {rect}")

                if not found:  # 첫 번째 것만 클릭
                    print("클릭 시도...")
                    from pywinauto import mouse
                    center_x = (rect.left + rect.right) // 2
                    center_y = (rect.top + rect.bottom) // 2
                    mouse.click(coords=(center_x, center_y))
                    print(f"좌표 ({center_x}, {center_y})로 클릭 완료!")
                    found = True
                    break
        except:
            pass

    if not found:
        print("'부양가족' 텍스트를 찾을 수 없습니다.")

except Exception as e:
    print(f"오류: {e}")
    import traceback
    traceback.print_exc()
