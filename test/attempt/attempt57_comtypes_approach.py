"""
시도 57: comtypes로 COM 인터페이스 접근

oleacc를 통해 IAccessible을 얻고,
거기서 텍스트 정보를 읽어오는 방식 시도
"""
import time


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 57: comtypes로 COM 접근")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt57_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 포커스 설정하고 현재 값 확인 (비교용)
        left_spread.set_focus()
        time.sleep(0.5)

        import pyperclip
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (type_keys로 복사): '{reference_value}'")

        print("\n=== 방법 1: UI Automation (UIA) 시도 ===")
        try:
            from pywinauto import uia

            # UIA 백엔드로 재연결 시도
            uia_app = uia.Application(backend='uia')
            uia_app.connect(title="사원등록")
            uia_dlg = uia_app.window(title="사원등록")

            print("✓ UIA 연결 성공")

            # UIA로 fpUSpread80 찾기
            try:
                # 여러 방법으로 시도
                uia_spread = uia_dlg.child_window(class_name="fpUSpread80", found_index=0)
                print(f"✓ UIA로 스프레드 찾음: {uia_spread}")

                # 텍스트 읽기 시도
                try:
                    text = uia_spread.window_text()
                    print(f"  window_text(): '{text}'")
                except Exception as e:
                    print(f"  window_text() 실패: {e}")

                # Value 패턴 시도
                try:
                    value = uia_spread.legacy_properties().Value
                    print(f"  legacy Value: '{value}'")
                except Exception as e:
                    print(f"  legacy Value 실패: {e}")

            except Exception as e:
                print(f"✗ UIA 스프레드 찾기 실패: {e}")

        except Exception as e:
            print(f"✗ UIA 방법 실패: {e}")

        print("\n=== 방법 2: IAccessible로 자식 요소 탐색 ===")
        try:
            from ctypes import windll, byref, c_int, c_void_p, POINTER
            from comtypes import GUID, IUnknown
            import comtypes.client

            # IAccessible GUID
            IID_IAccessible = GUID("{618736E0-3C3D-11CF-810C-00AA00389B71}")

            oleacc = windll.oleacc
            ptr = c_void_p()

            result = oleacc.AccessibleObjectFromWindow(
                hwnd,
                0,  # OBJID_CLIENT
                byref(IID_IAccessible),
                byref(ptr)
            )

            if result == 0 and ptr:
                print("✓ IAccessible 획득 성공")

                # QueryInterface로 제대로 된 COM 객체 얻기
                acc = comtypes.client.GetBestInterface(ptr)
                print(f"  타입: {type(acc)}")

                # 속성 시도
                try:
                    name = acc.accName(0)  # CHILDID_SELF = 0
                    print(f"  이름: '{name}'")
                except:
                    print("  이름 읽기 실패")

                try:
                    value = acc.accValue(0)
                    print(f"  값: '{value}'")
                except:
                    print("  값 읽기 실패")

                try:
                    role = acc.accRole(0)
                    print(f"  역할: {role}")
                except:
                    print("  역할 읽기 실패")

                # 자식 개수 확인
                try:
                    child_count = acc.accChildCount
                    print(f"  자식 개수: {child_count}")

                    if child_count > 0:
                        # 자식들 순회
                        children = acc.accChildren
                        print(f"  자식 탐색 중...")
                        for i, child in enumerate(children[:10]):  # 처음 10개만
                            try:
                                child_name = acc.accName(child)
                                child_value = acc.accValue(child)
                                print(f"    자식 {i}: 이름='{child_name}', 값='{child_value}'")
                            except:
                                pass

                except Exception as e:
                    print(f"  자식 탐색 실패: {e}")

            else:
                print(f"✗ IAccessible 획득 실패 (result={result})")

        except Exception as e:
            import traceback
            print(f"✗ IAccessible 방법 실패: {e}")
            traceback.print_exc()

        print("\n=== 방법 3: GetWindowLong으로 추가 정보 확인 ===")
        try:
            import win32api
            import win32con

            # 윈도우 정보
            style = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)
            ex_style = win32api.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            user_data = win32api.GetWindowLong(hwnd, win32con.GWL_USERDATA)

            print(f"스타일: 0x{style:08X}")
            print(f"확장 스타일: 0x{ex_style:08X}")
            print(f"사용자 데이터: 0x{user_data:08X}")

            # 클래스 정보
            class_name = win32api.GetClassName(hwnd)
            print(f"클래스명: '{class_name}'")

        except Exception as e:
            print(f"✗ GetWindowLong 실패: {e}")

        capture_func("attempt57_01_complete.png")

        return {
            "success": False,
            "message": "COM/UIA 접근 모두 실패 - 창 활성화 방식 필요"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
