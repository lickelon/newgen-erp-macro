"""
Attempt 12: fpUSpread80 COM 인터페이스로 접근

fpUSpread80은 Farpoint Spread ActiveX 컨트롤
→ COM/OLE 인터페이스로 데이터 입력 시도
"""
import time
import win32com.client
import pythoncom


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 12: fpUSpread80 COM 인터페이스")
    print("=" * 60)

    try:
        capture_func("attempt12_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt12_01_basic_tab.png")

        # fpUSpread80 찾기
        print("\n[2/4] fpUSpread80 컨트롤 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
                    hwnd = ctrl.handle
                    rect = ctrl.rectangle()
                    print(f"  발견: HWND=0x{hwnd:08X}")
                    print(f"        위치: ({rect.left},{rect.top}) 크기: {rect.width()}x{rect.height()}")
            except:
                pass

        if not spread_controls:
            return {"success": False, "message": "fpUSpread80 컨트롤 없음"}

        # 첫 번째 spread 사용 (기본사항 탭의 것)
        spread_ctrl = spread_controls[0]
        spread_hwnd = spread_ctrl.handle

        print(f"\n[3/4] COM 인터페이스 접근 시도...")

        # 방법 1: HWND로 COM 객체 얻기
        print("\n  방법 1: AccessibleObjectFromWindow")
        try:
            import ctypes
            from ctypes import POINTER, c_long
            from comtypes import IUnknown, GUID

            # IDispatch GUID
            IID_IDispatch = GUID("{00020400-0000-0000-C000-000000000046}")

            ole32 = ctypes.windll.ole32
            oleacc = ctypes.windll.oleacc

            ptr = POINTER(IUnknown)()
            result = ole32.AccessibleObjectFromWindow(
                spread_hwnd,
                0,  # OBJID_WINDOW
                ctypes.byref(IID_IDispatch),
                ctypes.byref(ptr)
            )

            if result == 0 and ptr:
                print("  ✓ COM 객체 획득 성공!")
                # TODO: COM 메서드 호출
            else:
                print(f"  ✗ 실패: result={result}")

        except Exception as e:
            print(f"  ✗ 오류: {e}")

        # 방법 2: pywinauto의 win32_element 사용
        print("\n  방법 2: pywinauto wrapper")
        try:
            # Spread 컨트롤 정보 출력
            print(f"  클래스명: {spread_ctrl.class_name()}")
            print(f"  컨트롤 ID: {spread_ctrl.control_id()}")

            # 하위 요소 확인
            children = spread_ctrl.children()
            print(f"  자식 요소: {len(children)}개")

        except Exception as e:
            print(f"  ✗ 오류: {e}")

        # 방법 3: 직접 메시지 전송
        print("\n  방법 3: 직접 윈도우 메시지")
        import win32api
        import win32con

        test_text = "TEST123"

        # 클릭 시뮬레이션 (셀 선택)
        try:
            print(f"    클릭 시뮬레이션...")
            x, y = 50, 30  # 첫 번째 셀 추정
            lparam = win32api.MAKELONG(x, y)
            win32api.SendMessage(spread_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32api.SendMessage(spread_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.2)

            print(f"    텍스트 입력: \"{test_text}\"")
            # WM_CHAR로 한 글자씩
            for char in test_text:
                win32api.SendMessage(spread_hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.02)

            # Enter 키
            time.sleep(0.1)
            win32api.SendMessage(spread_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            time.sleep(0.02)
            win32api.SendMessage(spread_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

            print(f"    ✓ 입력 완료")

        except Exception as e:
            print(f"    ✗ 오류: {e}")

        time.sleep(1)
        capture_func("attempt12_02_after_input.png")

        print("\n[4/4] 결과 확인")
        print("  스크린샷 attempt12_02_after_input.png 확인 필요")

        return {"success": True, "message": "fpUSpread80 접근 시도 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}


if __name__ == "__main__":
    from pywinauto import application
    import sys
    import os

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    result = run(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
