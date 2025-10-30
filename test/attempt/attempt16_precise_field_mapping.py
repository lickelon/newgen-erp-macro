"""
Attempt 16: 정밀 필드 매핑

Spread #0, #1의 다양한 좌표를 그리드로 테스트하여
사번, 성명, 주민번호 등 정확한 필드 위치 파악
"""
import time
import win32api
import win32con


def quick_input(hwnd, x, y, text):
    """빠른 입력 (간소화)"""
    try:
        lparam = win32api.MAKELONG(x, y)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.1)

        for char in text:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.01)

        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.01)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.15)
        return True
    except:
        return False


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 16: 정밀 필드 매핑 (Grid Search)")
    print("=" * 60)

    try:
        capture_func("attempt16_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)

        # fpUSpread 찾기
        print("\n[2/4] fpUSpread 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 2:
            return {"success": False, "message": "fpUSpread 부족"}

        spread0 = spread_controls[0]
        spread1 = spread_controls[1]

        print(f"  Spread #0: HWND=0x{spread0.handle:08X}")
        print(f"  Spread #1: HWND=0x{spread1.handle:08X}")

        # Grid 좌표 정의
        # X축: 왼쪽부터 오른쪽으로 여러 컬럼
        # Y축: 위쪽부터 아래쪽으로 여러 행
        x_positions = [50, 100, 150, 200, 250, 300, 350, 400]
        y_positions = [30, 50, 70, 90, 110]

        print("\n[3/4] Spread #1 그리드 테스트...")
        print(f"  X축: {x_positions}")
        print(f"  Y축: {y_positions}")

        test_data = [
            "A", "B", "C", "D", "E", "F", "G", "H",
        ]

        # Y=30 라인 테스트 (첫 번째 행)
        print("\n  Y=30 라인 (첫 번째 행):")
        for idx, x in enumerate(x_positions):
            if idx >= len(test_data):
                break
            text = test_data[idx]
            success = quick_input(spread1.handle, x, 30, text)
            status = "✓" if success else "✗"
            print(f"    X={x:3d}: \"{text}\" {status}")

        time.sleep(0.5)
        capture_func("attempt16_01_y30_line.png")

        # Y=50 라인 테스트 (두 번째 행)
        print("\n  Y=50 라인 (두 번째 행):")
        test_data2 = ["1", "2", "3", "4", "5", "6", "7", "8"]
        for idx, x in enumerate(x_positions):
            if idx >= len(test_data2):
                break
            text = test_data2[idx]
            success = quick_input(spread1.handle, x, 50, text)
            status = "✓" if success else "✗"
            print(f"    X={x:3d}: \"{text}\" {status}")

        time.sleep(0.5)
        capture_func("attempt16_02_y50_line.png")

        # Spread #0도 테스트
        print("\n  Spread #0 테스트:")
        test_coords = [
            (50, 30, "S0A"),
            (150, 30, "S0B"),
            (250, 30, "S0C"),
        ]

        for x, y, text in test_coords:
            success = quick_input(spread0.handle, x, y, text)
            status = "✓" if success else "✗"
            print(f"    ({x},{y}): \"{text}\" {status}")

        time.sleep(0.5)
        capture_func("attempt16_03_spread0.png")

        # 실제 데이터로 입력 테스트
        print("\n[4/4] 실제 데이터 입력 테스트...")

        # 주민번호가 들어간 위치 재시도 (분홍색으로 표시되었던 곳)
        # Spread #1의 특정 좌표
        real_tests = [
            (spread1.handle, 100, 30, "9999001", "사번추정1"),
            (spread1.handle, 200, 30, "홍길동", "성명추정1"),
            (spread1.handle, 300, 30, "910101-1234567", "주민번호추정1"),
        ]

        for hwnd, x, y, text, label in real_tests:
            success = quick_input(hwnd, x, y, text)
            status = "✓" if success else "✗"
            print(f"  [{label}] ({x},{y}): \"{text}\" {status}")

        time.sleep(1)
        capture_func("attempt16_04_final.png")

        return {"success": True, "message": "그리드 테스트 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}


if __name__ == "__main__":
    from pywinauto import application
    import sys
    import os

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    result = run(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
