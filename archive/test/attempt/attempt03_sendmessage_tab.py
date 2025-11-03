"""
시도 3: Win32 SendMessage로 탭 컨트롤에 직접 메시지 전송
마우스 커서를 움직이지 않고 윈도우 메시지만 사용
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
    print("시도 3: Win32 SendMessage로 탭 선택")
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
        print(f"탭 컨트롤 크기: W={rect.width()}, H={rect.height()}")

        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt03_00_initial.png")

        # 탭 높이 추정 (일반적으로 상단 25-35px)
        tab_height = 30

        # 탭 목록 (예상)
        # 탭1: 기본사항(현재), 탭2: 부양가족정보, 탭3: 소득자료 등
        # 각 탭의 대략적인 X 좌표 계산
        tab_width_estimate = 100  # 탭 하나당 대략적인 너비

        print("\n=== 방법 1: TCM_SETCURSEL 메시지로 탭 인덱스 변경 ===")

        TCM_SETCURSEL = 0x130C  # 탭 선택 메시지

        # 탭 인덱스 1, 2, 3 시도
        for tab_index in range(1, 4):
            print(f"\n탭 인덱스 {tab_index} 선택 시도...")
            result = win32api.SendMessage(hwnd, TCM_SETCURSEL, tab_index, 0)
            print(f"  SendMessage 결과: {result}")
            time.sleep(0.5)

            capture_func(f"attempt03_01_tcm_setcursel_{tab_index}.png")
            time.sleep(0.3)

        print("\n=== 방법 2: WM_LBUTTONDOWN/UP 메시지로 특정 좌표 클릭 ===")

        # 탭 영역의 각 위치를 클릭
        # 좌표는 탭 컨트롤 내부의 상대 좌표 (클라이언트 좌표)
        tab_y = 15  # 탭 영역의 중앙 Y 좌표

        click_positions = [
            (50, tab_y, "탭1 위치"),
            (150, tab_y, "탭2 위치"),
            (250, tab_y, "탭3 위치"),
            (350, tab_y, "탭4 위치"),
            (450, tab_y, "탭5 위치"),
        ]

        for i, (x, y, label) in enumerate(click_positions, 1):
            print(f"\n{label} ({x}, {y}) 클릭 시도...")

            # LPARAM = MAKELONG(x, y)
            lparam = win32api.MAKELONG(x, y)

            # WM_LBUTTONDOWN
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.1)

            # WM_LBUTTONUP
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.5)

            capture_func(f"attempt03_02_click_{i}_{label}.png")
            print(f"  ✓ 완료")

        print("\n=== 방법 3: WM_NOTIFY 메시지 시뮬레이션 ===")

        # 부모 윈도우에 탭 변경 알림 보내기
        parent_hwnd = win32api.GetParent(hwnd)
        print(f"\n부모 윈도우 HWND: {parent_hwnd}")

        # TCN_SELCHANGE 알림 (탭 선택 변경)
        TCN_SELCHANGE = -551  # (TCN_FIRST - 1)

        for tab_index in range(1, 3):
            print(f"\n부모에게 탭 {tab_index} 선택 알림...")
            # 먼저 탭 선택
            win32api.SendMessage(hwnd, 0x130C, tab_index, 0)  # TCM_SETCURSEL
            time.sleep(0.2)

            # 부모에게 알림 (실제 구조체가 필요하지만 간단히 시도)
            # WM_NOTIFY 메시지는 복잡해서 작동 안 할 수 있음
            time.sleep(0.5)

            capture_func(f"attempt03_03_notify_{tab_index}.png")

        return {"success": True, "message": "모든 방법 시도 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
