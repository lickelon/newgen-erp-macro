"""
시도 54: fpUSpread80에서 클립보드 복사 테스트

여러 방법으로 클립보드 복사를 시도하고 어떤 방법이 작동하는지 확인
"""
import time
import win32api
import win32con
import win32gui
import pyperclip


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 54: fpUSpread80 클립보드 복사 테스트")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt54_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")
        print(f"위치: {left_spread.rectangle()}")

        # 스프레드에 포커스 설정
        left_spread.set_focus()
        time.sleep(0.5)
        capture_func("attempt54_01_focused.png")

        # 클립보드 초기화
        pyperclip.copy("BEFORE_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 1: WM_COPY 메시지 ===")
        WM_COPY = 0x0301
        win32api.SendMessage(hwnd, WM_COPY, 0, 0)
        time.sleep(0.3)
        result1 = pyperclip.paste()
        print(f"결과: '{result1}'")
        success1 = result1 != "BEFORE_TEST"
        print(f"{'✓' if success1 else '✗'} WM_COPY: {'성공' if success1 else '실패'}")

        # 클립보드 초기화
        pyperclip.copy("BEFORE_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 2: Ctrl+C (SendMessage) ===")
        # Ctrl 누름
        lparam_ctrl_down = 1 | (0x1D << 16)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, lparam_ctrl_down)
        time.sleep(0.02)
        # C 누름
        lparam_c_down = 1 | (0x2E << 16)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('C'), lparam_c_down)
        time.sleep(0.02)
        # C 뗌
        lparam_c_up = 1 | (0x2E << 16) | (1 << 30) | (1 << 31)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, ord('C'), lparam_c_up)
        time.sleep(0.02)
        # Ctrl 뗌
        lparam_ctrl_up = 1 | (0x1D << 16) | (1 << 30) | (1 << 31)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, lparam_ctrl_up)
        time.sleep(0.3)
        result2 = pyperclip.paste()
        print(f"결과: '{result2}'")
        success2 = result2 != "BEFORE_TEST"
        print(f"{'✓' if success2 else '✗'} Ctrl+C (SendMessage): {'성공' if success2 else '실패'}")

        # 클립보드 초기화
        pyperclip.copy("BEFORE_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 3: left_spread.type_keys('^c', set_foreground=False) ===")
        left_spread.type_keys("^c", pause=0.05, set_foreground=False)
        time.sleep(0.3)
        result3 = pyperclip.paste()
        print(f"결과: '{result3}'")
        success3 = result3 != "BEFORE_TEST"
        print(f"{'✓' if success3 else '✗'} type_keys (set_foreground=False): {'성공' if success3 else '실패'}")

        # 클립보드 초기화
        pyperclip.copy("BEFORE_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 4: left_spread.type_keys('^c', set_foreground=True) ===")
        left_spread.type_keys("^c", pause=0.05, set_foreground=True)
        time.sleep(0.3)
        result4 = pyperclip.paste()
        print(f"결과: '{result4}'")
        success4 = result4 != "BEFORE_TEST"
        print(f"{'✓' if success4 else '✗'} type_keys (set_foreground=True): {'성공' if success4 else '실패'}")

        # 클립보드 초기화
        pyperclip.copy("BEFORE_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 5: dlg.type_keys('^c') ===")
        dlg.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        result5 = pyperclip.paste()
        print(f"결과: '{result5}'")
        success5 = result5 != "BEFORE_TEST"
        print(f"{'✓' if success5 else '✗'} dlg.type_keys: {'성공' if success5 else '실패'}")

        capture_func("attempt54_02_tests_complete.png")

        # 요약
        print("\n" + "="*60)
        print("요약:")
        print(f"  1. WM_COPY: {'✓' if success1 else '✗'}")
        print(f"  2. SendMessage Ctrl+C: {'✓' if success2 else '✗'}")
        print(f"  3. type_keys (set_foreground=False): {'✓' if success3 else '✗'}")
        print(f"  4. type_keys (set_foreground=True): {'✓' if success4 else '✗'}")
        print(f"  5. dlg.type_keys: {'✓' if success5 else '✗'}")
        print("="*60)

        # 어떤 방법이든 하나라도 성공하면 성공
        any_success = success1 or success2 or success3 or success4 or success5

        if any_success:
            working_methods = []
            if success1: working_methods.append("WM_COPY")
            if success2: working_methods.append("SendMessage Ctrl+C")
            if success3: working_methods.append("type_keys (no foreground)")
            if success4: working_methods.append("type_keys (foreground)")
            if success5: working_methods.append("dlg.type_keys")

            return {
                "success": True,
                "message": f"성공한 방법: {', '.join(working_methods)}"
            }
        else:
            return {
                "success": False,
                "message": "모든 복사 방법 실패"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
