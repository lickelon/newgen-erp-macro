"""
Attempt 40: SendInput으로 실제 키보드 입력 시뮬레이션

포커스된 컨트롤에 직접 키 입력 (전역 키보드 이벤트)
"""
import time
import ctypes
from ctypes import wintypes


# SendInput 구조체 정의
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))
    ]


class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]
    _anonymous_ = ("_input",)
    _fields_ = [("type", wintypes.DWORD), ("_input", _INPUT)]


KEYEVENTF_UNICODE = 0x0004
INPUT_KEYBOARD = 1


def send_unicode_char(char):
    """유니코드 문자를 SendInput으로 전송"""
    inputs = []
    for c in char:
        # Key down
        ki_down = KEYBDINPUT(0, ord(c), KEYEVENTF_UNICODE, 0, None)
        inputs.append(INPUT(INPUT_KEYBOARD, INPUT._INPUT(ki=ki_down)))
        # Key up
        ki_up = KEYBDINPUT(0, ord(c), KEYEVENTF_UNICODE | 0x0002, 0, None)
        inputs.append(INPUT(INPUT_KEYBOARD, INPUT._INPUT(ki=ki_up)))

    array_type = INPUT * len(inputs)
    input_array = array_type(*inputs)
    ctypes.windll.user32.SendInput(len(inputs), ctypes.byref(input_array), ctypes.sizeof(INPUT))


def send_key(vk_code):
    """가상 키 코드를 SendInput으로 전송"""
    # Key down
    ki_down = KEYBDINPUT(vk_code, 0, 0, 0, None)
    input_down = INPUT(INPUT_KEYBOARD, INPUT._INPUT(ki=ki_down))

    # Key up
    ki_up = KEYBDINPUT(vk_code, 0, 0x0002, 0, None)
    input_up = INPUT(INPUT_KEYBOARD, INPUT._INPUT(ki=ki_up))

    ctypes.windll.user32.SendInput(1, ctypes.byref(input_down), ctypes.sizeof(INPUT))
    time.sleep(0.02)
    ctypes.windll.user32.SendInput(1, ctypes.byref(input_up), ctypes.sizeof(INPUT))


def test_sendinput(dlg, capture_func):
    """SendInput 방식 테스트"""
    print("\n" + "=" * 60)
    print("Attempt 40: SendInput 방식")
    print("=" * 60)

    try:
        from tab_automation import TabAutomation
        import win32con

        # 현재 상태
        print("\n[1/6] 현재 상태...")
        capture_func("attempt40_00_initial.png")
        time.sleep(0.5)

        tab_auto = TabAutomation()
        tab_auto.connect()

        # 기본사항 탭
        print("\n[2/6] 기본사항 탭 선택...")
        tab_auto.select_tab("기본사항")
        time.sleep(1.0)
        capture_func("attempt40_01_basic_tab.png")

        # 부양가족정보 탭
        print("\n[3/6] 부양가족정보 탭 선택...")
        tab_auto.select_tab("부양가족정보")
        time.sleep(1.0)
        capture_func("attempt40_02_dependents_tab.png")

        # Down 키
        print("\n[4/6] Down 키로 행 이동...")
        send_key(win32con.VK_DOWN)
        time.sleep(0.5)
        capture_func("attempt40_03_after_down.png")

        # 데이터 입력 (SendInput 방식)
        print("\n[5/6] SendInput으로 데이터 입력...")
        test_data = ["4", "최부양", "내", "2018"]
        field_names = ["연말관계", "성명", "내외국", "년도"]

        for idx, value in enumerate(test_data):
            print(f"  [{idx+1}/4] {field_names[idx]}: \"{value}\"")

            # 문자 입력
            send_unicode_char(value)
            time.sleep(0.3)
            print(f"    → 입력 완료")

            # Tab (마지막 필드 제외)
            if idx < len(test_data) - 1:
                send_key(win32con.VK_TAB)
                time.sleep(0.2)
                print(f"    → Tab")

            capture_func(f"attempt40_04_field{idx+1}.png")

        # Enter로 확정
        print("\n[6/6] Enter로 확정...")
        send_key(win32con.VK_RETURN)
        time.sleep(0.5)

        capture_func("attempt40_05_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "SendInput 방식 완료",
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

    result = test_sendinput(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
