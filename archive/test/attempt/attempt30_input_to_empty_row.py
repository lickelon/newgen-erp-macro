"""
Attempt 30: 기존 빈 행에 더블클릭 후 입력

이미 추가된 빈 행 중 하나를 더블클릭하여 입력
"""
import time
import win32api
import win32con


def input_to_existing_empty_row(dlg, capture_func):
    """기존 빈 행에 입력"""
    print("\n" + "=" * 60)
    print("Attempt 30: 빈 행에 입력")
    print("=" * 60)

    try:
        capture_func("attempt30_00_initial.png")

        # 부양가족정보 탭 선택
        print("\n[1/4] 부양가족정보 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("부양가족정보")
        time.sleep(0.5)
        capture_func("attempt30_01_dependents_tab.png")

        # fpUSpread80 찾기
        print("\n[2/4] 스프레드시트 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) == 0:
            return {"success": False, "message": "스프레드시트 없음"}

        spread = spread_controls[0]
        hwnd = spread.handle
        print(f"  Spread #0 HWND: 0x{hwnd:08X}")

        # 두 번째 빈 행(본인 행 다음)에 더블클릭
        print("\n[3/4] 두 번째 행 (첫 번째 빈 행) 더블클릭...")

        # y=30: 본인 행 (0번)
        # y=60: 두 번째 행 (첫 번째 부양가족)
        cell_x = 100  # 성명 컬럼
        cell_y = 60   # 두 번째 행

        lparam = win32api.MAKELONG(cell_x, cell_y)
        print(f"  좌표: ({cell_x}, {cell_y})")

        # 더블클릭
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.3)

        capture_func("attempt30_02_after_doubleclick.png")

        # 성명부터 입력 (성명 컬럼에서 시작)
        print("\n[4/4] 데이터 입력...")
        test_data = ["박부양", "자녀", "2020", "5"]

        for idx, value in enumerate(test_data):
            print(f"  필드 {idx+1}: \"{value}\"")

            # 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.2)
            capture_func(f"attempt30_03_field{idx+1}.png")

            # Tab으로 다음 필드
            if idx < len(test_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)

        # Enter로 확정
        print("\n  Enter로 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt30_04_final.png")

        print("\n" + "=" * 60)
        print("완료: 빈 행에 입력 시도")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "빈 행 입력 완료",
            "data": test_data,
            "coordinates": f"({cell_x}, {cell_y})"
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

    result = input_to_existing_empty_row(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
