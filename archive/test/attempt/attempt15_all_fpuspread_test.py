"""
Attempt 15: 모든 fpUSpread80에 순차적으로 입력 테스트

3개의 fpUSpread가 있는데, 각각의 역할을 파악하고
각각에 입력을 시도해봅니다.
"""
import time
import win32api
import win32con


def test_spread_input(hwnd, spread_name, test_cells):
    """fpUSpread에 여러 좌표로 입력 시도"""
    print(f"\n  {spread_name} 테스트:")
    print(f"  HWND=0x{hwnd:08X}")

    success_count = 0
    for cell_info in test_cells:
        x, y, text, label = cell_info

        try:
            # 클릭
            lparam = win32api.MAKELONG(x, y)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.03)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.15)

            # 입력
            for char in text:
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            # Enter
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            time.sleep(0.02)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            time.sleep(0.2)

            success_count += 1
            print(f"    [{label}] ({x},{y}) \"{text}\" ✓")

        except Exception as e:
            print(f"    [{label}] ({x},{y}) ✗ {e}")

    return success_count


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 15: 모든 fpUSpread80 테스트")
    print("=" * 60)

    try:
        capture_func("attempt15_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt15_01_basic_tab.png")

        # 모든 fpUSpread80 찾기
        print("\n[2/4] 모든 fpUSpread80 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    rect = ctrl.rectangle()
                    spread_controls.append({
                        "ctrl": ctrl,
                        "hwnd": ctrl.handle,
                        "rect": rect
                    })
            except:
                pass

        print(f"  총 {len(spread_controls)}개 발견:")
        for idx, spread in enumerate(spread_controls):
            rect = spread["rect"]
            print(f"  [{idx}] HWND=0x{spread['hwnd']:08X}")
            print(f"       위치: ({rect.left},{rect.top})")
            print(f"       크기: {rect.width()}x{rect.height()}")

        if len(spread_controls) < 2:
            return {"success": False, "message": "fpUSpread80 부족"}

        # 각 스프레드에 테스트 데이터 입력
        print("\n[3/4] 각 fpUSpread에 입력 시도...")

        # 공통 테스트 셀 좌표
        test_cells_pattern1 = [
            (50, 30, "A001", "A1"),
            (150, 30, "이름A", "B1"),
            (250, 30, "900101", "C1"),
        ]

        test_cells_pattern2 = [
            (100, 50, "B002", "A2"),
            (200, 50, "이름B", "B2"),
            (300, 50, "900202", "C2"),
        ]

        results = []

        # Spread 0 (첫 번째)
        if len(spread_controls) > 0:
            result = test_spread_input(
                spread_controls[0]["hwnd"],
                "Spread #0",
                test_cells_pattern1
            )
            results.append(("Spread #0", result))

        time.sleep(1)
        capture_func("attempt15_02_after_spread0.png")

        # Spread 1 (두 번째)
        if len(spread_controls) > 1:
            result = test_spread_input(
                spread_controls[1]["hwnd"],
                "Spread #1",
                test_cells_pattern2
            )
            results.append(("Spread #1", result))

        time.sleep(1)
        capture_func("attempt15_03_after_spread1.png")

        # Spread 2 (세 번째)
        if len(spread_controls) > 2:
            result = test_spread_input(
                spread_controls[2]["hwnd"],
                "Spread #2",
                test_cells_pattern1
            )
            results.append(("Spread #2", result))

        time.sleep(1)
        capture_func("attempt15_04_after_spread2.png")

        # 결과 요약
        print("\n[4/4] 결과 요약:")
        for spread_name, count in results:
            print(f"  {spread_name}: {count}/3개 입력")

        return {
            "success": True,
            "message": f"총 {len(results)}개 Spread 테스트 완료"
        }

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
