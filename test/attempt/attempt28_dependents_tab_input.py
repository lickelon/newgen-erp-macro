"""
Attempt 28: 부양가족정보 탭의 스프레드시트에 입력

부양가족정보 탭 선택 → 스프레드시트 찾기 → 입력
"""
import time
import win32api
import win32con


def input_to_dependents_tab(dlg, capture_func):
    """부양가족정보 탭의 스프레드시트에 입력"""
    print("\n" + "=" * 60)
    print("Attempt 28: 부양가족정보 탭 입력")
    print("=" * 60)

    try:
        capture_func("attempt28_00_initial.png")

        # 부양가족정보 탭 선택
        print("\n[1/4] 부양가족정보 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("부양가족정보")
        time.sleep(0.5)
        capture_func("attempt28_01_dependents_tab.png")

        # fpUSpread80 찾기
        print("\n[2/4] 스프레드시트 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        print(f"  총 {len(spread_controls)}개 fpUSpread80 발견")

        if len(spread_controls) == 0:
            return {"success": False, "message": "스프레드시트 없음"}

        # 첫 번째 스프레드 사용
        spread = spread_controls[0]
        hwnd = spread.handle
        print(f"  Spread #0 HWND: 0x{hwnd:08X}")

        # 더블클릭으로 첫 셀 선택
        print("\n[3/4] 첫 번째 셀 더블클릭...")
        cell_x = 50
        cell_y = 30

        lparam = win32api.MAKELONG(cell_x, cell_y)

        # 더블클릭
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.3)

        capture_func("attempt28_02_after_doubleclick.png")

        # 테스트 데이터 입력
        print("\n[4/4] 데이터 입력...")
        test_data = ["홍길동", "자녀", "2015", "10"]  # 이름, 관계, 출생년도, 나이

        for idx, value in enumerate(test_data):
            print(f"  필드 {idx+1}: \"{value}\"")

            # 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.2)

            # Tab으로 다음 필드
            if idx < len(test_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)

            capture_func(f"attempt28_03_field{idx+1}.png")

        # Enter로 확정
        print("\n  Enter로 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt28_04_final.png")

        print("\n" + "=" * 60)
        print("완료: 부양가족정보 탭에 입력 시도")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "부양가족정보 탭 입력 완료",
            "data": test_data,
            "spread_count": len(spread_controls)
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

    result = input_to_dependents_tab(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
