"""
Attempt 38: 탭 선택 진단 정보 수집

탭 컨트롤 정보를 상세히 출력하여 문제 진단
"""
import time


def test_tab_diagnostic(dlg, capture_func):
    """탭 선택 진단"""
    print("\n" + "=" * 60)
    print("Attempt 38: 탭 선택 진단")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/5] 현재 상태...")
        capture_func("attempt38_00_initial.png")
        time.sleep(0.5)

        # 모든 탭 컨트롤 찾기
        print("\n[2/5] 탭 컨트롤 검색 중...")
        tab_controls = []
        for ctrl in dlg.descendants():
            try:
                class_name = ctrl.class_name()
                if class_name.startswith("Afx:TabWnd:"):
                    tab_controls.append(ctrl)
                    print(f"  발견: {class_name}")
                    print(f"    HWND: 0x{ctrl.handle:08X}")
                    print(f"    Visible: {ctrl.is_visible()}")
                    print(f"    Enabled: {ctrl.is_enabled()}")
                    try:
                        rect = ctrl.rectangle()
                        print(f"    Rect: {rect}")
                    except:
                        print(f"    Rect: (error)")
            except:
                pass

        print(f"\n  → 총 {len(tab_controls)}개의 탭 컨트롤 발견")

        if len(tab_controls) == 0:
            return {"success": False, "message": "탭 컨트롤 없음"}

        # 첫 번째 탭 컨트롤 사용
        tab_ctrl = tab_controls[0]
        hwnd = tab_ctrl.handle
        print(f"\n  → 사용할 탭 컨트롤: 0x{hwnd:08X}")

        # 탭 선택 테스트 1: 기본사항
        print("\n[3/5] 기본사항 탭 선택 시도...")
        import win32api
        import win32con

        x, y = 50, 15
        lparam = win32api.MAKELONG(x, y)

        print(f"  클릭 좌표: ({x}, {y})")
        print(f"  LPARAM: 0x{lparam:08X}")

        result_down = win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        print(f"  WM_LBUTTONDOWN 결과: {result_down}")
        time.sleep(0.1)

        result_up = win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        print(f"  WM_LBUTTONUP 결과: {result_up}")
        time.sleep(1.0)

        capture_func("attempt38_01_after_basic.png")
        print("  → 스크린샷 저장")

        # 탭 선택 테스트 2: 부양가족정보
        print("\n[4/5] 부양가족정보 탭 선택 시도...")
        x, y = 150, 15
        lparam = win32api.MAKELONG(x, y)

        print(f"  클릭 좌표: ({x}, {y})")
        print(f"  LPARAM: 0x{lparam:08X}")

        result_down = win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        print(f"  WM_LBUTTONDOWN 결과: {result_down}")
        time.sleep(0.1)

        result_up = win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        print(f"  WM_LBUTTONUP 결과: {result_up}")
        time.sleep(1.0)

        capture_func("attempt38_02_after_dependents.png")
        print("  → 스크린샷 저장")

        # 다시 기본사항
        print("\n[5/5] 다시 기본사항 탭 선택 시도...")
        x, y = 50, 15
        lparam = win32api.MAKELONG(x, y)

        result_down = win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        print(f"  WM_LBUTTONDOWN 결과: {result_down}")
        time.sleep(0.1)

        result_up = win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        print(f"  WM_LBUTTONUP 결과: {result_up}")
        time.sleep(1.0)

        capture_func("attempt38_03_after_basic_again.png")
        print("  → 스크린샷 저장")

        print("\n" + "=" * 60)
        print("진단 완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "진단 완료",
            "tab_controls_found": len(tab_controls)
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }


if __name__ == "__main__":
    from pywinauto import application
    import sys
    import os

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    result = test_tab_diagnostic(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
