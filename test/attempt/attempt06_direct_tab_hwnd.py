"""
시도 6: 탭 컨트롤 HWND를 직접 찾아서 WM_LBUTTONDOWN 전송
Spy++ 스크린샷에서 확인한 탭 컨트롤에 직접 메시지 전송
"""
import time
import win32api
import win32con
import win32gui

def run(dlg, capture_func):
    """
    Args:
        dlg: pywinauto 윈도우 객체
        capture_func: 스크린샷 함수 (filename) -> None

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 6: 탭 컨트롤에 직접 WM_LBUTTONDOWN 전송")
    print("="*60)

    try:
        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt06_00_initial.png")

        # 탭 컨트롤 찾기
        tab_control = dlg.child_window(class_name="Afx:TabWnd:cd0000:8:10003:10", found_index=0)

        if not tab_control.exists():
            return {"success": False, "message": "탭 컨트롤을 찾을 수 없음"}

        tab_hwnd = tab_control.handle
        rect = tab_control.rectangle()

        print(f"\n탭 컨트롤 HWND: 0x{tab_hwnd:08X}")
        print(f"탭 컨트롤 위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")
        print(f"탭 컨트롤 크기: W={rect.width()}, H={rect.height()}")

        # 탭 위치들 (클라이언트 좌표)
        tab_positions = [
            (50, 15, "기본사항"),
            (150, 15, "부양가족정보"),
            (250, 15, "소득자료"),
        ]

        print("\n=== 탭 컨트롤에 직접 WM_LBUTTONDOWN/UP 전송 ===")

        for i, (x, y, tab_name) in enumerate(tab_positions, 1):
            print(f"\n[{i}] '{tab_name}' 탭 ({x}, {y}) 클릭 시도...")

            # LPARAM = MAKELONG(x, y)
            lparam = win32api.MAKELONG(x, y)

            print(f"  탭 컨트롤 HWND: 0x{tab_hwnd:08X}")
            print(f"  좌표: ({x}, {y})")
            print(f"  LPARAM: 0x{lparam:08X}")

            # WM_LBUTTONDOWN 전송
            result = win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            print(f"  WM_LBUTTONDOWN 결과: {result}")
            time.sleep(0.1)

            # WM_LBUTTONUP 전송
            result = win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            print(f"  WM_LBUTTONUP 결과: {result}")
            time.sleep(0.5)

            # 스크린샷
            capture_func(f"attempt06_01_tab_{i}_{tab_name}.png")
            print(f"  ✓ 완료")
            time.sleep(0.5)

        print("\n=== PostMessage로 시도 (SendMessage 대신) ===")

        # PostMessage는 메시지를 큐에 넣고 즉시 리턴
        for i, (x, y, tab_name) in enumerate(tab_positions[:2], 1):
            print(f"\n[{i}] '{tab_name}' 탭 PostMessage 시도...")

            lparam = win32api.MAKELONG(x, y)

            # PostMessage 사용
            win32gui.PostMessage(tab_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32gui.PostMessage(tab_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.5)

            capture_func(f"attempt06_02_post_{i}_{tab_name}.png")
            print(f"  ✓ 완료")

        return {"success": True, "message": "모든 시도 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
