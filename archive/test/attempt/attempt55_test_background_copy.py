"""
시도 55: 백그라운드(비활성화) 상태에서 클립보드 복사 테스트

다른 창을 활성화한 상태에서 fpUSpread80의 복사가 작동하는지 확인
"""
import time
import win32api
import win32con
import win32gui
import pyperclip
import subprocess


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 55: 백그라운드 상태에서 클립보드 복사 테스트")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt55_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 스프레드에 포커스 설정 (초기 값 확인용)
        left_spread.set_focus()
        time.sleep(0.5)

        # 초기 값 확인
        pyperclip.copy("BEFORE")
        time.sleep(0.2)
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        initial_value = pyperclip.paste()
        print(f"사원등록 창 활성화 상태에서 복사한 값: '{initial_value}'")

        # 메모장 실행하여 다른 창 활성화
        print("\n메모장을 실행하여 사원등록 창을 비활성화합니다...")
        notepad = subprocess.Popen(['notepad.exe'])
        time.sleep(2)  # 메모장이 완전히 뜰 때까지 대기

        capture_func("attempt55_01_notepad_opened.png")

        # 현재 활성 창 확인
        active_hwnd = win32gui.GetForegroundWindow()
        active_title = win32gui.GetWindowText(active_hwnd)
        print(f"현재 활성 창: '{active_title}' (HWND: 0x{active_hwnd:08X})")

        if "사원등록" in active_title:
            print("⚠️  사원등록 창이 여전히 활성화되어 있습니다!")
        else:
            print("✓ 사원등록 창이 비활성화되었습니다")

        # 클립보드 초기화
        pyperclip.copy("BACKGROUND_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 1: SendMessage Ctrl+C (백그라운드) ===")
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
        time.sleep(0.5)

        result1 = pyperclip.paste()
        print(f"결과: '{result1}'")
        success1 = (result1 != "BACKGROUND_TEST" and result1 == initial_value)
        print(f"{'✓' if success1 else '✗'} SendMessage Ctrl+C: {'성공' if success1 else '실패'}")
        if not success1 and result1 != "BACKGROUND_TEST":
            print(f"  (다른 값이 복사됨 - 메모장에서 복사된 것일 수 있음)")

        # 클립보드 초기화
        pyperclip.copy("BACKGROUND_TEST")
        time.sleep(0.2)

        print("\n=== 테스트 2: left_spread.type_keys (백그라운드) ===")
        left_spread.type_keys("^c", pause=0.05, set_foreground=False)
        time.sleep(0.5)

        result2 = pyperclip.paste()
        print(f"결과: '{result2}'")
        success2 = (result2 != "BACKGROUND_TEST" and result2 == initial_value)
        print(f"{'✓' if success2 else '✗'} type_keys (set_foreground=False): {'성공' if success2 else '실패'}")
        if not success2 and result2 != "BACKGROUND_TEST":
            print(f"  (다른 값이 복사됨 - 메모장에서 복사된 것일 수 있음)")

        capture_func("attempt55_02_tests_complete.png")

        # 메모장 종료
        print("\n메모장을 종료합니다...")
        notepad.terminate()
        time.sleep(1)

        # 요약
        print("\n" + "="*60)
        print("요약:")
        print(f"  초기 값 (활성화 상태): '{initial_value}'")
        print(f"  1. SendMessage Ctrl+C (백그라운드): {'✓' if success1 else '✗'} -> '{result1}'")
        print(f"  2. type_keys (백그라운드): {'✓' if success2 else '✗'} -> '{result2}'")
        print("="*60)

        if success1 or success2:
            working_methods = []
            if success1: working_methods.append("SendMessage Ctrl+C")
            if success2: working_methods.append("type_keys")

            return {
                "success": True,
                "message": f"백그라운드 복사 성공: {', '.join(working_methods)}"
            }
        else:
            return {
                "success": False,
                "message": "백그라운드 상태에서 복사 불가능"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
