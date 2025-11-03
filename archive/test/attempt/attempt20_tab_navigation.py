"""
Attempt 20: Tab 키 네비게이션으로 정확한 셀 이동

좌표 대신 Tab 키로 셀 간 이동하며 입력
더 안정적이고 예측 가능함
"""
import time
import win32api
import win32con


def input_with_tab_navigation(dlg, capture_func):
    """Tab 키로 셀 이동하며 입력"""
    print("\n" + "=" * 60)
    print("Attempt 20: Tab 키 네비게이션")
    print("=" * 60)

    try:
        capture_func("attempt20_00_initial.png")

        # fpUSpread80 찾기
        print("\n[1/5] fpUSpread80 찾기...")
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

        # 첫 번째 셀 클릭 (좌상단)
        print("\n[2/5] 첫 번째 셀 클릭...")
        lparam = win32api.MAKELONG(10, 10)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.3)
        capture_func("attempt20_01_first_cell.png")

        # 직원 데이터
        employee_data = ["2025200", "탭테스트", "930101-1234567", "32"]

        print("\n[3/5] Tab으로 이동하며 입력...")
        for idx, value in enumerate(employee_data):
            print(f"  필드 {idx+1}: \"{value}\"")

            # 현재 셀에 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.2)

            # Tab 키로 다음 셀로 이동
            if idx < len(employee_data) - 1:  # 마지막 필드가 아니면
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)
                print(f"    → Tab으로 다음 셀 이동")

            capture_func(f"attempt20_02_field{idx+1}.png")

        # Enter로 확정
        print("\n[4/5] Enter로 행 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt20_03_final.png")

        print("\n[5/5] 완료")
        print("  ✓ Tab 키로 4개 필드 이동하며 입력")
        print("  ⚠️  스크린샷을 확인하여 올바른 위치에 입력되었는지 검증 필요")

        return {
            "success": True,
            "message": "Tab 네비게이션으로 입력 완료",
            "data": employee_data
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

    result = input_with_tab_navigation(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
