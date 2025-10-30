"""
Attempt 35: 기본사항 → 부양가족정보 → 아래 방향키 → 입력

탭 전환으로 (1,1) 선택 → Down 키로 아래 행 이동 → 입력
"""
import time
import win32api
import win32con


def input_after_arrow_down(dlg, capture_func):
    """탭 전환 후 방향키로 이동하여 입력"""
    print("\n" + "=" * 60)
    print("Attempt 35: 탭 전환 + 방향키 이동 + 입력")
    print("=" * 60)

    try:
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()

        # 기본사항 탭 선택
        print("\n[1/5] 기본사항 탭 선택...")
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt35_01_basic_tab.png")

        # 부양가족정보 탭 선택
        print("\n[2/5] 부양가족정보 탭 선택...")
        tab_auto.select_tab("부양가족정보")
        time.sleep(0.5)
        capture_func("attempt35_02_dependents_tab.png")
        print("  → (1,1) 셀 선택됨")

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

        # 아래 방향키 입력
        print("\n[3/5] 아래 방향키 입력...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.5)
        capture_func("attempt35_03_after_arrow_down.png")
        print("  → 아래 행으로 이동")

        # 데이터 입력
        print("\n[4/5] 데이터 입력...")
        test_data = [
            "4",          # 연말관계: 4=자녀
            "김자녀",      # 성명
            "내",         # 내/외국
            "2010",       # 출생년도
        ]

        print(f"  입력 데이터: {test_data}")

        for idx, value in enumerate(test_data):
            field_names = ["연말관계", "성명", "내외국", "출생년도"]
            print(f"    [{idx+1}] {field_names[idx]}: \"{value}\"")

            # WM_CHAR로 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.3)
            capture_func(f"attempt35_04_field{idx+1}.png")

            # Tab으로 다음 필드
            if idx < len(test_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)
                print(f"      → Tab")

        # Enter로 확정
        print("\n[5/5] Enter로 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt35_05_final.png")

        print("\n" + "=" * 60)
        print("완료: 탭 전환 + 방향키 이동 + 입력")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "방향키 이동 후 입력 완료",
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

    result = input_after_arrow_down(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
