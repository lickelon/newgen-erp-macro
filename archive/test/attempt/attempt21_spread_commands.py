"""
Attempt 21: fpUSpread80 SendCommand/PostCommand 메서드 활용

Attempt 19에서 발견한 SendCommand, PostCommand를 실제로 사용
Farpoint Spread는 특수 명령어로 제어 가능할 수 있음
"""
import time
import win32api
import win32con
import win32gui


def explore_spread_methods(dlg, capture_func):
    """SendCommand/PostCommand 메서드 탐색 및 사용"""
    print("\n" + "=" * 60)
    print("Attempt 21: fpUSpread 명령어 방식 제어")
    print("=" * 60)

    try:
        capture_func("attempt21_00_initial.png")

        # 기본사항 탭
        print("\n[1/6] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt21_01_basic_tab.png")

        # fpUSpread80 찾기
        print("\n[2/6] fpUSpread80 찾기...")
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

        # 컨트롤의 모든 메서드 탐색
        print("\n[3/6] 사용 가능한 메서드 탐색...")
        all_methods = [attr for attr in dir(left_list)
                      if not attr.startswith('_') and callable(getattr(left_list, attr))]

        interesting_methods = [m for m in all_methods
                             if any(keyword in m.lower()
                                   for keyword in ['send', 'post', 'set', 'get',
                                                  'text', 'cell', 'row', 'col',
                                                  'click', 'select', 'edit'])]

        print(f"  관심 메서드 ({len(interesting_methods)}개):")
        for method in interesting_methods[:20]:  # 처음 20개만
            print(f"    - {method}")

        # SendMessage로 특수 명령 시도
        print("\n[4/6] fpSpread 특수 메시지 탐색...")

        # Farpoint Spread의 알려진 메시지들 시도
        # SSM_SETTEXT, SSM_GETTEXT 등

        # WM_USER 기반 메시지들 (일반적으로 WM_USER + N)
        print("  WM_USER 기반 메시지 시도:")

        test_messages = [
            (0x0400 + 0,   "WM_USER+0"),
            (0x0400 + 1,   "WM_USER+1"),
            (0x0400 + 10,  "WM_USER+10 (SetActiveCell?)"),
            (0x0400 + 11,  "WM_USER+11 (GetActiveCell?)"),
            (0x0400 + 20,  "WM_USER+20 (SetText?)"),
            (0x0400 + 21,  "WM_USER+21 (GetText?)"),
            (0x0400 + 50,  "WM_USER+50 (AddRow?)"),
            (0x0400 + 100, "WM_USER+100"),
            (0x0400 + 200, "WM_USER+200"),
        ]

        results = {}
        for msg, desc in test_messages:
            try:
                result = win32api.SendMessage(hwnd, msg, 0, 0)
                results[desc] = result
                if result != 0:
                    print(f"    {desc}: {result} ✓")
            except Exception as e:
                results[desc] = f"오류: {e}"

        # pywinauto의 SendCommand/PostCommand 시도
        print("\n[5/6] SendCommand/PostCommand 메서드 시도...")

        # 시도 1: 다양한 명령 ID 테스트
        print("  시도 1: 명령 ID 테스트...")
        for cmd_id in [0, 1, 2, 10, 100, 1000]:
            try:
                result = left_list.send_command(cmd_id)
                print(f"    send_command({cmd_id}): {result}")
            except Exception as e:
                print(f"    send_command({cmd_id}): {e}")
                break  # 하나 실패하면 중단

        # 시도 2: 특수 윈도우 메시지 조합
        print("\n  시도 2: 윈도우 메시지로 셀 선택 시도...")

        # 셀 선택을 위한 메시지 조합
        # 일부 그리드 컨트롤은 LVM_SETITEMSTATE 같은 메시지 사용
        LVM_FIRST = 0x1000
        LVM_SETITEMSTATE = LVM_FIRST + 43

        try:
            # 첫 번째 행 선택 시도
            result = win32api.SendMessage(hwnd, LVM_SETITEMSTATE, 0, 0)
            print(f"    LVM_SETITEMSTATE(0): {result}")
        except Exception as e:
            print(f"    LVM_SETITEMSTATE 실패: {e}")

        # 시도 3: 키보드 입력으로 셀 이동 후 입력
        print("\n  시도 3: Home 키로 첫 셀 이동 후 입력...")

        # Home 키로 첫 셀로 이동
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.2)

        capture_func("attempt21_02_after_home.png")

        # Down 키로 데이터 행으로 이동 (헤더 건너뛰기)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.2)

        capture_func("attempt21_03_after_down.png")

        # 직원 데이터 입력
        employee_data = ["2025300", "명령테스트", "940101-1234567", "31"]

        print("\n[6/6] 키보드 네비게이션으로 입력...")
        for idx, value in enumerate(employee_data):
            print(f"  필드 {idx+1}: \"{value}\"")

            # 현재 셀에 입력
            for char in value:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            time.sleep(0.2)
            capture_func(f"attempt21_04_field{idx+1}.png")

            # Tab으로 다음 셀로 이동
            if idx < len(employee_data) - 1:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)

        # Enter로 확정
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        capture_func("attempt21_05_final.png")

        print("\n" + "=" * 60)
        print("완료: Home → Down → 입력 → Tab 방식 시도")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": "키보드 네비게이션 완료",
            "messages_tested": len(test_messages),
            "methods_found": len(interesting_methods)
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

    result = explore_spread_methods(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
