"""
시도 5: 키보드 입력으로 탭 이동
SendKeys를 사용하여 마우스 움직임 없이 탭 이동
"""
import time
import win32api
import win32con

def run(dlg, capture_func):
    """
    Args:
        dlg: pywinauto 윈도우 객체
        capture_func: 스크린샷 함수 (filename) -> None

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 5: 키보드 입력으로 탭 이동")
    print("="*60)

    try:
        # 탭 컨트롤 찾기
        tab_control = dlg.child_window(class_name="Afx:TabWnd:cd0000:8:10003:10", found_index=0)

        if not tab_control.exists():
            return {"success": False, "message": "탭 컨트롤을 찾을 수 없음"}

        rect = tab_control.rectangle()
        hwnd = tab_control.handle

        print(f"\n탭 컨트롤 HWND: {hwnd}")
        print(f"탭 컨트롤 위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")

        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt05_00_initial.png")

        print("\n=== 방법 1: WM_KEYDOWN으로 화살표 키 전송 ===")

        # 탭 컨트롤에 포커스 설정
        print("\n탭 컨트롤에 포커스 설정...")
        win32api.SetFocus(hwnd)
        time.sleep(0.3)

        # 오른쪽 화살표 키로 탭 이동 (VK_RIGHT)
        for i in range(1, 4):
            print(f"\n{i}번째 화살표 키 전송...")

            # WM_KEYDOWN: VK_RIGHT
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
            time.sleep(0.1)

            # WM_KEYUP: VK_RIGHT
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
            time.sleep(0.5)

            capture_func(f"attempt05_01_arrow_right_{i}.png")
            print(f"  ✓ 완료")

        print("\n=== 방법 2: Ctrl+Tab으로 탭 이동 ===")

        # 첫 번째 탭으로 이동 (왼쪽 화살표 3번)
        print("\n첫 번째 탭으로 리셋...")
        for i in range(3):
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_LEFT, 0)
            time.sleep(0.1)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_LEFT, 0)
            time.sleep(0.2)

        time.sleep(0.5)
        capture_func("attempt05_02_reset.png")

        # Ctrl+Tab으로 탭 이동
        for i in range(1, 3):
            print(f"\n{i}번째 Ctrl+Tab 전송...")

            # WM_KEYDOWN: VK_CONTROL
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            time.sleep(0.05)

            # WM_KEYDOWN: VK_TAB
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
            time.sleep(0.05)

            # WM_KEYUP: VK_TAB
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
            time.sleep(0.05)

            # WM_KEYUP: VK_CONTROL
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.5)

            capture_func(f"attempt05_03_ctrl_tab_{i}.png")
            print(f"  ✓ 완료")

        return {"success": True, "message": "모든 키보드 입력 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
