"""
시도 58: UI Automation 제대로 시도

pywinauto의 UIA 백엔드를 올바르게 사용하여
백그라운드에서도 값을 읽을 수 있는지 테스트
"""
import time
import subprocess


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 58: UI Automation 제대로 시도")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt58_00_initial.png")

        # 왼쪽 스프레드 찾기 (win32 백엔드)
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 참조 값 확인 (win32로 복사)
        import pyperclip
        left_spread.set_focus()
        time.sleep(0.5)
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (win32 백엔드): '{reference_value}'")

        print("\n=== 방법 1: Application(backend='uia')로 재연결 ===")
        try:
            from pywinauto import Application

            # UIA 백엔드로 새로 연결
            uia_app = Application(backend='uia')
            uia_app.connect(title="사원등록")
            uia_dlg = uia_app.window(title="사원등록")

            print("✓ UIA 백엔드로 연결 성공")

            # 모든 컨트롤 출력
            print("\nUIA 컨트롤 목록:")
            try:
                uia_dlg.print_control_identifiers(depth=3, filename=None)
            except:
                pass

            # fpUSpread80 찾기
            try:
                # class_name으로 찾기
                uia_spreads = uia_dlg.children(class_name="fpUSpread80")
                print(f"\nfpUSpread80 개수: {len(uia_spreads)}")

                if uia_spreads:
                    uia_spreads.sort(key=lambda s: s.rectangle().left)
                    uia_left_spread = uia_spreads[0]
                    print(f"✓ UIA로 왼쪽 스프레드 찾음")
                    print(f"  HWND: 0x{uia_left_spread.handle:08X}")
                    print(f"  위치: {uia_left_spread.rectangle()}")

                    # 여러 방법으로 텍스트 읽기 시도
                    print("\n텍스트 읽기 시도:")

                    try:
                        text = uia_left_spread.window_text()
                        print(f"  1. window_text(): '{text}'")
                    except Exception as e:
                        print(f"  1. window_text() 실패: {e}")

                    try:
                        texts = uia_left_spread.texts()
                        print(f"  2. texts(): {texts}")
                    except Exception as e:
                        print(f"  2. texts() 실패: {e}")

                    try:
                        # 자식 컨트롤 탐색
                        children = uia_left_spread.children()
                        print(f"  3. 자식 컨트롤 수: {len(children)}")
                        for i, child in enumerate(children[:10]):
                            try:
                                child_text = child.window_text()
                                if child_text:
                                    print(f"     자식 {i}: '{child_text}'")
                            except:
                                pass
                    except Exception as e:
                        print(f"  3. 자식 탐색 실패: {e}")

                    # UIA 요소로 직접 접근
                    try:
                        elem = uia_left_spread.element_info
                        print(f"  4. element_info: {elem}")

                        # UIA 속성 읽기
                        name = elem.name
                        automation_id = elem.automation_id
                        control_type = elem.control_type

                        print(f"     이름: '{name}'")
                        print(f"     AutomationId: '{automation_id}'")
                        print(f"     ControlType: '{control_type}'")

                    except Exception as e:
                        print(f"  4. element_info 실패: {e}")

                    # Value Pattern 시도
                    try:
                        value_pattern = uia_left_spread.wrapper_object()
                        print(f"  5. wrapper_object: {value_pattern}")

                        # Legacy Pattern
                        if hasattr(value_pattern, 'legacy_properties'):
                            props = value_pattern.legacy_properties()
                            print(f"     legacy_properties: {props}")

                        # Value
                        if hasattr(value_pattern, 'GetCurrentPropertyValue'):
                            from pywinauto.uia_defines import UIA_ValueValuePropertyId
                            value = value_pattern.GetCurrentPropertyValue(UIA_ValueValuePropertyId)
                            print(f"     Value: '{value}'")

                    except Exception as e:
                        print(f"  5. Value Pattern 실패: {e}")

                else:
                    print("✗ UIA로 fpUSpread80을 찾지 못함")

            except Exception as e:
                import traceback
                print(f"✗ fpUSpread80 찾기 실패: {e}")
                traceback.print_exc()

        except Exception as e:
            import traceback
            print(f"✗ UIA 백엔드 연결 실패: {e}")
            traceback.print_exc()

        print("\n=== 방법 2: 메모장으로 백그라운드 테스트 ===")
        try:
            print("메모장 실행하여 사원등록 창 비활성화...")
            notepad = subprocess.Popen(['notepad.exe'])
            time.sleep(2)

            import win32gui
            active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            print(f"현재 활성 창: '{active_title}'")

            # 백그라운드에서 UIA로 텍스트 읽기 시도
            try:
                if 'uia_left_spread' in locals():
                    bg_text = uia_left_spread.window_text()
                    print(f"✓ 백그라운드에서 읽기 성공: '{bg_text}'")
                else:
                    print("✗ uia_left_spread가 없음")
            except Exception as e:
                print(f"✗ 백그라운드 읽기 실패: {e}")

            # 메모장 종료
            notepad.terminate()
            time.sleep(0.5)

        except Exception as e:
            print(f"✗ 백그라운드 테스트 실패: {e}")

        capture_func("attempt58_01_complete.png")

        return {
            "success": False,
            "message": "UIA로도 셀 값 직접 읽기 불가능"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
