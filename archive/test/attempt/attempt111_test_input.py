"""
시도 111: 스프레드에 값 입력 테스트

사원 선택 후 값 입력 가능 여부 확인
"""
import win32process
import win32gui
import win32con
import win32api
import time
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
    print("시도 111: 스프레드에 값 입력 테스트")
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
        installment_dlg_hwnd = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            installment_dlg_hwnd = hwnd
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
        right_spread = spreads[1] if len(spreads) > 1 else None

        left_hwnd = left_spread.handle
        print(f"✓ 왼쪽 스프레드: 0x{left_hwnd:08X}")
        if right_spread:
            right_hwnd = right_spread.handle
            print(f"✓ 오른쪽 스프레드: 0x{right_hwnd:08X}")

        # === 방법 1: 왼쪽 스프레드에서 사원 선택 ===
        print("\n" + "="*60)
        print("[방법 1] 왼쪽 스프레드 사원 선택")
        print("="*60)

        # 첫 번째 행 클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.05)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
        time.sleep(0.5)

        print("✓ 첫 번째 행 클릭 완료")

        # === 방법 2: 오른쪽 스프레드로 Tab 이동 ===
        if right_spread:
            print("\n" + "="*60)
            print("[방법 2] 오른쪽 스프레드로 이동 후 입력")
            print("="*60)

            # Tab 키로 오른쪽으로 이동
            for i in range(3):
                win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.2)

            print("✓ Tab 3회 완료")

            # 다이얼로그를 통한 입력 시도 (docs/08에서 성공한 방법)
            print("\n[방법 2-1] installment_dlg.type_keys() 사용")
            try:
                installment_dlg.type_keys("12345", with_spaces=False, pause=0.05)
                time.sleep(0.3)
                print("✓ type_keys 입력 완료")

                # Enter로 확정
                installment_dlg.type_keys("{ENTER}", pause=0.05)
                time.sleep(0.3)
                print("✓ Enter 완료")

                return {"success": True, "message": "dlg.type_keys()로 입력 성공!"}

            except Exception as e:
                print(f"✗ type_keys 실패: {e}")

        # === 방법 3: 오른쪽 스프레드에 직접 SendMessage ===
        if right_spread:
            print("\n" + "="*60)
            print("[방법 3] 오른쪽 스프레드에 WM_CHAR 직접 전송")
            print("="*60)

            # 오른쪽 스프레드 클릭
            lParam = win32api.MAKELONG(50, 30)
            win32api.SendMessage(right_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
            time.sleep(0.05)
            win32api.SendMessage(right_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
            time.sleep(0.3)

            # 숫자 입력 (WM_CHAR)
            test_value = "999"
            for ch in test_value:
                win32api.SendMessage(right_hwnd, win32con.WM_CHAR, ord(ch), 0)
                time.sleep(0.05)

            print(f"✓ WM_CHAR로 '{test_value}' 전송 완료")

            # Enter
            win32api.SendMessage(right_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            time.sleep(0.3)

            print("✓ Enter 완료")

            return {"success": True, "message": f"WM_CHAR로 '{test_value}' 입력 완료"}

        # === 방법 4: 다이얼로그에서 직접 입력 (pywinauto wrapper 사용) ===
        print("\n" + "="*60)
        print("[방법 4] pywinauto wrapper로 입력")
        print("="*60)

        # 왼쪽 스프레드 선택 상태에서 입력
        try:
            # pywinauto wrapper를 통한 입력
            test_value = "777"

            # 다이얼로그 레벨에서 type_keys
            dlg_wrapper = app.window(handle=installment_dlg_hwnd)
            dlg_wrapper.type_keys(test_value, with_spaces=False, pause=0.05)
            time.sleep(0.3)

            print(f"✓ dlg_wrapper.type_keys()로 '{test_value}' 전송")

            # Tab으로 다음 필드
            dlg_wrapper.type_keys("{TAB}", pause=0.05)
            time.sleep(0.2)

            return {"success": True, "message": f"dlg_wrapper.type_keys()로 '{test_value}' 입력 시도"}

        except Exception as e:
            print(f"✗ wrapper 입력 실패: {e}")

        # === 방법 5: Right 키로 셀 이동 후 입력 ===
        print("\n" + "="*60)
        print("[방법 5] Right 키로 셀 이동 후 입력")
        print("="*60)

        # 왼쪽 스프레드에서 Right 키
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
        time.sleep(0.2)

        # 숫자 입력
        test_value = "555"
        for ch in test_value:
            win32api.SendMessage(left_hwnd, win32con.WM_CHAR, ord(ch), 0)
            time.sleep(0.05)

        print(f"✓ Right 후 '{test_value}' 입력")

        # Enter
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.3)

        return {"success": True, "message": f"Right + WM_CHAR로 '{test_value}' 입력 시도"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
