"""
Attempt 22: 왼쪽 스프레드시트에 명시적 포커스 + 키보드 네비게이션

좌표 사용 안 함. 대신:
1. SetFocus를 스프레드시트 HWND에 직접 호출
2. Ctrl+Home으로 A1 셀로 이동
3. Arrow 키로 원하는 위치 이동
4. 입력
"""
import time
import win32api
import win32con
import win32gui


def input_to_left_spread(dlg, capture_func):
    """왼쪽 스프레드시트에 포커스 후 키보드로 입력"""
    print("\n" + "=" * 60)
    print("Attempt 22: 왼쪽 스프레드 포커스 + 키보드")
    print("=" * 60)

    try:
        capture_func("attempt22_00_initial.png")

        # 기본사항 탭
        print("\n[1/5] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt22_01_basic_tab.png")

        # fpUSpread80 찾기
        print("\n[2/5] 왼쪽 스프레드 찾기...")
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

        # 명시적으로 포커스 설정 시도
        print("\n[3/5] 왼쪽 스프레드에 포커스...")

        # 방법 1: WM_SETFOCUS 메시지
        try:
            win32api.SendMessage(hwnd, win32con.WM_SETFOCUS, 0, 0)
            print("  WM_SETFOCUS 전송 완료")
            time.sleep(0.2)
        except Exception as e:
            print(f"  WM_SETFOCUS 실패: {e}")

        # 방법 2: 스프레드시트 내부 클릭 (중앙 안전 위치)
        # 좌표를 최소한으로 사용 - 그냥 중앙 클릭
        try:
            rect = left_list.rectangle()
            center_x = (rect.right - rect.left) // 2
            center_y = (rect.bottom - rect.top) // 2

            # 클라이언트 좌표로 중앙 클릭
            lparam = win32api.MAKELONG(center_x, center_y)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.03)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.3)
            print(f"  중앙 클릭: ({center_x}, {center_y})")
        except Exception as e:
            print(f"  중앙 클릭 실패: {e}")

        capture_func("attempt22_02_after_focus.png")

        # Ctrl+Home으로 A1 셀로 이동
        print("\n[4/5] Ctrl+Home으로 A1 이동...")

        # Ctrl 누름
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        time.sleep(0.05)

        # Home 누름
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.05)

        # Ctrl 뗌
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        print("  Ctrl+Home 전송 완료")
        capture_func("attempt22_03_after_ctrlhome.png")

        # 첫 번째 데이터 행으로 이동 (Down 키)
        print("  Down 키로 첫 데이터 행 이동...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.3)

        capture_func("attempt22_04_after_down.png")

        # 직원 데이터 입력
        employee_data = [
            ("사번", "2025400"),
            ("성명", "포커스테스트"),
            ("주민번호", "950101-1234567"),
            ("나이", "30")
        ]

        print("\n[5/5] 데이터 입력 (Tab으로 이동)...")
        for idx, (field_name, value) in enumerate(employee_data):
            print(f"  {field_name}: \"{value}\"")

            # 현재 셀에 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.2)
            capture_func(f"attempt22_05_field{idx+1}_{field_name}.png")

            # Tab으로 다음 셀 이동
            if idx < len(employee_data) - 1:
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

        capture_func("attempt22_06_final.png")

        print("\n" + "=" * 60)
        print("완료: 포커스 → Ctrl+Home → Down → 입력")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "왼쪽 스프레드 포커스 + 키보드 입력 완료",
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

    result = input_to_left_spread(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
