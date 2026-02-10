"""
시도 104: 인접 컬럼 이동하여 읽기 시도

사원코드 컬럼이 선택 불가이므로, 인접한 컬럼으로 이동하여
복사 가능한 컬럼이 있는지 확인
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
    print("시도 104: 인접 컬럼 이동하여 읽기 시도")
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
        print(f"왼쪽 스프레드: 0x{hwnd:08X}")

        # 3. 포커스 설정 및 Home 키로 시작 위치로
        print("\n[3단계] 초기 위치 설정")

        left_spread.set_focus()
        time.sleep(0.3)

        # Home 키로 첫 셀로 이동
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.2)

        print("✓ Home 키로 첫 셀 이동")

        # 4. 여러 컬럼을 오른쪽으로 이동하며 복사 시도
        print("\n[4단계] 컬럼별 복사 시도")

        results = []

        for col_idx in range(5):  # 5개 컬럼 시도
            print(f"\n[컬럼 {col_idx}]")

            # 복사 시도
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
                results.append((col_idx, value))
            else:
                print(f"  ✗ 클립보드 비어있음 (복사 불가 컬럼)")

            # 다음 컬럼으로 이동 (마지막 시도 제외)
            if col_idx < 4:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
                time.sleep(0.2)

        # 5. 결과 정리
        print("\n[5단계] 결과 정리")

        if results:
            print(f"\n✓ {len(results)}개 컬럼에서 데이터 읽기 성공:")
            for col_idx, value in results:
                print(f"  컬럼 {col_idx}: '{value}'")

            return {
                "success": True,
                "message": f"{len(results)}개 컬럼 읽기 성공: {results}"
            }
        else:
            return {
                "success": False,
                "message": "모든 컬럼에서 복사 실패 - 다른 방법 필요"
            }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
