"""
시도 101: 분납적용 다이얼로그에서 사원코드 읽기
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
    print("시도 101: 분납적용 다이얼로그에서 사원코드 읽기")
    print("="*60)

    try:
        # 1. 급여자료입력 연결
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[1단계] 급여자료입력 연결 (PID: {process_id})")

        # 2. 분납적용 다이얼로그 찾기
        print("\n[2단계] 분납적용 다이얼로그 찾기")

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
                        if "분납적용" in text or "분납(환급)계산" in text:
                            installment_dlg = dialog
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 3. fpUSpread80 스프레드 찾기
        print("\n[3단계] 스프레드 컨트롤 찾기")

        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        print(f"✓ {len(spreads)}개 스프레드 발견")

        if len(spreads) < 2:
            return {"success": False, "message": f"스프레드가 충분하지 않습니다 ({len(spreads)}개)"}

        # 왼쪽 스프레드 (사원 목록)
        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"\n왼쪽 스프레드 (사원 목록):")
        print(f"  HWND: 0x{left_spread.handle:08X}")
        print(f"  위치: {left_spread.rectangle()}")

        # 4. 사원코드 읽기 시도
        print("\n[4단계] 사원코드 읽기 시도")

        hwnd = left_spread.handle

        # 방법 1: 포커스 + 클립보드 복사
        print("\n[방법 1] 포커스 + Ctrl+C")

        try:
            # 포커스 설정
            left_spread.set_focus()
            time.sleep(0.3)

            # Home 키로 첫 셀로 이동
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
            time.sleep(0.2)

            # 클립보드 초기화
            pyperclip.copy("")

            # Ctrl+C
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, ord('C'), 0)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.3)

            value = pyperclip.paste()
            print(f"  복사된 값: '{value}'")

            if value:
                print(f"  ✓ 성공! 값: '{value}'")

                # Right 키로 다음 셀 이동
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
                time.sleep(0.2)

                # 다음 셀 복사
                pyperclip.copy("")
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, ord('C'), 0)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
                time.sleep(0.3)

                value2 = pyperclip.paste()
                print(f"  다음 셀 값: '{value2}'")

                return {"success": True, "message": f"사원코드 읽기 성공: 첫 번째='{value}', 두 번째='{value2}'"}
            else:
                print("  ✗ 클립보드에 값이 없습니다")

        except Exception as e:
            print(f"  ✗ 오류: {e}")

        return {"success": False, "message": "사원코드를 읽지 못했습니다"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
