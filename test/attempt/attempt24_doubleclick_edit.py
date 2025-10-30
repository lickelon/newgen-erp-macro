"""
Attempt 24: 더블클릭으로 셀 편집 모드 진입

일반 스프레드시트처럼 더블클릭으로 셀을 활성화한 후 입력
"""
import time
import win32api
import win32con


def modify_with_doubleclick(dlg, capture_func):
    """더블클릭으로 셀 편집 모드 진입 후 입력"""
    print("\n" + "=" * 60)
    print("Attempt 24: 더블클릭 셀 편집")
    print("=" * 60)

    try:
        capture_func("attempt24_00_initial.png")

        # 기본사항 탭
        print("\n[1/3] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt24_01_basic_tab.png")

        # 왼쪽 스프레드 찾기
        print("\n[2/3] 왼쪽 스프레드 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            return {"success": False, "message": "Spread 부족"}

        left_list = spread_controls[2]
        hwnd = left_list.handle
        print(f"  Spread #2 HWND: 0x{hwnd:08X}")

        # 첫 번째 데이터 셀 더블클릭 (234번 직원, 사번 컬럼)
        print("\n[3/3] 첫 번째 데이터 셀 더블클릭 후 입력...")

        # 추정 좌표 (사번 컬럼, 첫 데이터 행)
        # Attempt 15에서 성공한 좌표 사용
        cell_x = 50
        cell_y = 30

        lparam = win32api.MAKELONG(cell_x, cell_y)

        # 더블클릭 시뮬레이션
        print(f"  ({cell_x}, {cell_y}) 더블클릭...")

        # 첫 번째 클릭
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.05)

        # 두 번째 클릭 (더블클릭)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.3)

        capture_func("attempt24_02_after_doubleclick.png")

        # 입력 (사번)
        test_value = "DBL001"
        print(f"  입력: \"{test_value}\"")

        for char in test_value:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.015)

        time.sleep(0.3)
        capture_func("attempt24_03_after_input.png")

        # Enter로 확정
        print("  Enter...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt24_04_final.png")

        print("\n" + "=" * 60)
        print("완료: 더블클릭 → 입력 → Enter")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "더블클릭 입력 완료",
            "value": test_value
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

    result = modify_with_doubleclick(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
