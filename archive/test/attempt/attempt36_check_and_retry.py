"""
Attempt 36: 현재 상태 확인 후 정확하게 재시도

1. 현재 상태 스크린샷
2. 기본사항 탭 강제 선택
3. 부양가족정보 탭 선택
4. Down 키
5. 입력
"""
import time
import win32api
import win32con


def check_and_retry_input(dlg, capture_func):
    """현재 상태 확인 후 재시도"""
    print("\n" + "=" * 60)
    print("Attempt 36: 상태 확인 후 재시도")
    print("=" * 60)

    try:
        # 현재 상태 확인
        print("\n[0/6] 현재 상태 확인...")
        capture_func("attempt36_00_current_state.png")
        time.sleep(0.5)

        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()

        # 기본사항 탭 강제 선택
        print("\n[1/6] 기본사항 탭 강제 선택...")
        tab_auto.select_tab("기본사항")
        time.sleep(1.0)  # 충분히 대기
        capture_func("attempt36_01_basic_tab.png")
        print("  → 기본사항 탭 선택됨")

        # 부양가족정보 탭 선택
        print("\n[2/6] 부양가족정보 탭 선택...")
        tab_auto.select_tab("부양가족정보")
        time.sleep(1.0)  # 충분히 대기
        capture_func("attempt36_02_dependents_tab.png")
        print("  → 부양가족정보 탭 선택됨")

        # 스프레드 찾기
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 1:
            return {"success": False, "message": "스프레드 없음"}

        spread = spread_controls[0]
        hwnd = spread.handle
        print(f"  Spread HWND: 0x{hwnd:08X}")

        # Down 키 입력
        print("\n[3/6] Down 키 입력...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.5)
        capture_func("attempt36_03_after_down.png")
        print("  → Down 키 전송 완료")

        # 데이터 입력
        print("\n[4/6] 데이터 입력 시작...")
        test_data = ["4", "박자녀", "내", "2015"]

        print(f"  입력 데이터: {test_data}")

        for idx, value in enumerate(test_data):
            field_names = ["연말관계", "성명", "내외국", "년도"]
            print(f"    [{idx+1}] {field_names[idx]}: \"{value}\"")

            # 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.3)

            # Tab
            if idx < len(test_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)
                print(f"      → Tab")

            capture_func(f"attempt36_04_field{idx+1}.png")

        # Enter
        print("\n[5/6] Enter로 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt36_05_final.png")

        print("\n[6/6] 완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "재시도 완료",
            "data": test_data
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

    result = check_and_retry_input(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
