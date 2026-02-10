"""
시도 102: 셀 클릭 후 사원코드 읽기
"""
import win32process
import win32gui
import win32con
import win32api
import time
import pyperclip
from pywinauto import Application


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 102: 셀 클릭 후 사원코드 읽기")
    print("="*60)

    try:
        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[1단계] 급여자료입력 연결 (PID: {process_id})")

        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "#32770":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        installment_dlg = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 2. 스프레드 찾기
        print("\n[2단계] 스프레드 찾기")

        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        hwnd = left_spread.handle
        rect = left_spread.rectangle()
        print(f"왼쪽 스프레드: 0x{hwnd:08X}")
        print(f"  위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")
        print(f"  크기: {rect.width()} x {rect.height()}")

        # 3. 셀 클릭 시도
        print("\n[3단계] 셀 클릭 시도")

        # 여러 위치 시도
        test_positions = [
            (30, 30, "좌상단"),
            (50, 30, "사원코드 예상 위치"),
            (100, 30, "중앙"),
            (30, 50, "아래쪽"),
        ]

        for x, y, desc in test_positions:
            print(f"\n[시도] {desc} ({x}, {y})")

            # WM_LBUTTONDOWN
            lParam = win32api.MAKELONG(x, y)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
            time.sleep(0.1)

            # WM_LBUTTONUP
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lParam)
            time.sleep(0.3)

            # 클립보드 복사 시도
            pyperclip.copy("")

            # Ctrl+C
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, ord('C'), 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.3)

            value = pyperclip.paste()

            if value:
                print(f"  ✓ 성공! 값: '{value}'")
                return {"success": True, "message": f"사원코드 읽기 성공: '{value}' at ({x}, {y})"}
            else:
                print(f"  ✗ 클립보드 비어있음")

        # 4. Down 키로 다른 행 시도
        print("\n[4단계] Down 키로 다른 행 이동 후 시도")

        # 첫 번째 셀 클릭
        lParam = win32api.MAKELONG(30, 30)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lParam)
        time.sleep(0.2)

        # Down 키
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.3)

        # 복사 시도
        pyperclip.copy("")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, ord('C'), 0)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        value = pyperclip.paste()

        if value:
            print(f"  ✓ Down 후 성공! 값: '{value}'")
            return {"success": True, "message": f"Down 후 읽기 성공: '{value}'"}
        else:
            print(f"  ✗ 여전히 클립보드 비어있음")

        return {"success": False, "message": "모든 방법으로 사원코드를 읽지 못했습니다"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
