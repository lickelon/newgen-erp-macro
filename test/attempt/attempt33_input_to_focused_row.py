"""
Attempt 33: 현재 포커스된 행에 바로 입력

아무것도 클릭하지 않고, 현재 포커스된 상태에서 바로 입력
"""
import time
import win32api
import win32con


def input_to_focused_row(dlg, capture_func):
    """현재 포커스된 행에 바로 입력"""
    print("\n" + "=" * 60)
    print("Attempt 33: 포커스된 행에 바로 입력")
    print("=" * 60)

    try:
        # 초기 상태 스크린샷
        capture_func("attempt33_00_initial.png")
        print("\n현재 상태에서 바로 입력 시작...")

        # 스프레드 찾기 (SendMessage용)
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 1:
            return {"success": False, "message": "스프레드 없음"}

        # 첫 번째 스프레드 (오른쪽 위)
        spread = spread_controls[0]
        hwnd = spread.handle

        # 테스트 데이터 - 부양가족 정보
        test_data = ["최부양", "배우자", "1985", "40"]

        print(f"\n입력할 데이터: {test_data}")

        for idx, value in enumerate(test_data):
            print(f"  필드 {idx+1}: \"{value}\"")

            # WM_CHAR로 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.2)
            capture_func(f"attempt33_01_field{idx+1}.png")

            # Tab으로 다음 필드
            if idx < len(test_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)
                print(f"    → Tab")

        # Enter로 확정
        print("\n  Enter로 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt33_02_final.png")

        print("\n" + "=" * 60)
        print("완료: 포커스된 행에 직접 입력")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "포커스된 행에 입력 완료",
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

    result = input_to_focused_row(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
