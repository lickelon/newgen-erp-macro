"""
Attempt 42: pywinauto의 type_keys 메서드 사용

pywinauto가 제공하는 고수준 메서드 사용
"""
import time


def test_type_keys(dlg, capture_func):
    """pywinauto type_keys 테스트"""
    print("\n" + "=" * 60)
    print("Attempt 42: pywinauto type_keys")
    print("=" * 60)

    try:
        from tab_automation import TabAutomation
        import win32con
        import win32api

        # 현재 상태
        print("\n[1/6] 현재 상태...")
        capture_func("attempt42_00_initial.png")
        time.sleep(0.5)

        tab_auto = TabAutomation()
        tab_auto.connect()

        # 기본사항 탭
        print("\n[2/6] 기본사항 탭 선택...")
        tab_auto.select_tab("기본사항")
        time.sleep(1.0)
        capture_func("attempt42_01_basic_tab.png")

        # 부양가족정보 탭
        print("\n[3/6] 부양가족정보 탭 선택...")
        tab_auto.select_tab("부양가족정보")
        time.sleep(1.0)
        capture_func("attempt42_02_dependents_tab.png")

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
        print(f"  Spread HWND: 0x{spread.handle:08X}")

        # Down 키 (SendMessage 사용)
        print("\n[4/6] Down 키로 행 이동...")
        win32api.SendMessage(spread.handle, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(spread.handle, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.5)
        capture_func("attempt42_03_after_down.png")

        # 데이터 입력 (type_keys 사용)
        print("\n[5/6] type_keys로 데이터 입력...")
        test_data = ["4", "정부양", "내", "2017"]
        field_names = ["연말관계", "성명", "내외국", "년도"]

        for idx, value in enumerate(test_data):
            print(f"  [{idx+1}/4] {field_names[idx]}: \"{value}\"")

            # type_keys 사용
            spread.type_keys(value, with_spaces=False, pause=0.05)
            time.sleep(0.3)
            print(f"    → 입력 완료")

            # Tab (마지막 필드 제외)
            if idx < len(test_data) - 1:
                spread.type_keys("{TAB}", pause=0.05)
                time.sleep(0.2)
                print(f"    → Tab")

            capture_func(f"attempt42_04_field{idx+1}.png")

        # Enter로 확정
        print("\n[6/6] Enter로 확정...")
        spread.type_keys("{ENTER}", pause=0.05)
        time.sleep(0.5)

        capture_func("attempt42_05_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "type_keys 방식 완료",
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

    result = test_type_keys(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
