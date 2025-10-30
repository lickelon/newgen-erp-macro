"""
Attempt 34: 연말관계 코드를 먼저 입력

연말관계 값(숫자)을 먼저 입력해야 행이 활성화됨
"""
import time
import win32api
import win32con


def input_with_relationship_code(dlg, capture_func):
    """연말관계 코드부터 입력"""
    print("\n" + "=" * 60)
    print("Attempt 34: 연말관계 코드로 입력")
    print("=" * 60)

    try:
        # 초기 상태 스크린샷
        capture_func("attempt34_00_initial.png")
        print("\n연말관계 코드(3=배우자)부터 입력...")

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

        # 첫 번째 스프레드 (오른쪽 위 부양가족)
        spread = spread_controls[0]
        hwnd = spread.handle

        # 테스트 데이터 - 배우자
        # 연말관계: 3=배우자
        test_data = [
            "3",          # 연말관계 코드
            "최부양",      # 성명
            "내",         # 내/외국
            "1985",       # 출생년도 또는 다른 값
        ]

        print(f"\n입력할 데이터: {test_data}")
        print(f"  연말관계 코드 3 = 배우자")

        for idx, value in enumerate(test_data):
            field_names = ["연말관계", "성명", "내외국", "기타"]
            field_name = field_names[idx]
            print(f"  [{idx+1}] {field_name}: \"{value}\"")

            # WM_CHAR로 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.3)
            capture_func(f"attempt34_01_field{idx+1}_{field_name}.png")

            # Tab으로 다음 필드
            if idx < len(test_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.3)
                print(f"    → Tab")

        # Enter로 확정
        print("\n  Enter로 확정...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt34_02_final.png")

        print("\n" + "=" * 60)
        print("완료: 연말관계 코드로 입력")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "연말관계 코드 입력 완료",
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

    result = input_with_relationship_code(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
