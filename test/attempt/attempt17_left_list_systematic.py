"""
Attempt 17: 왼쪽 사원 목록(Spread #2) 체계적 테스트

각 컬럼을 하나씩 테스트하고 매번 스크린샷 촬영
목표: 사번, 성명, 주민번호, 나이 컬럼의 정확한 좌표 파악
"""
import time
import win32api
import win32con


def single_cell_input(hwnd, x, y, text, capture_func, screenshot_name):
    """단일 셀 입력 + 스크린샷"""
    print(f"  좌표 ({x},{y}): \"{text}\"")

    try:
        # 클릭
        lparam = win32api.MAKELONG(x, y)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.2)

        # 입력
        for char in text:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.015)

        # Enter
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.3)

        # 스크린샷
        capture_func(screenshot_name)
        print(f"  ✓ 스크린샷: {screenshot_name}")

        return True

    except Exception as e:
        print(f"  ✗ 오류: {e}")
        return False


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 17: 왼쪽 사원 목록 체계적 테스트")
    print("=" * 60)

    try:
        capture_func("attempt17_00_initial.png")

        # 기본사항 탭
        print("\n[1/3] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt17_01_basic_tab.png")

        # Spread #2 찾기 (왼쪽 목록)
        print("\n[2/3] Spread #2 (왼쪽 목록) 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            return {"success": False, "message": "Spread #2 없음"}

        left_list = spread_controls[2]  # 세 번째 = 왼쪽 목록
        hwnd = left_list.handle
        rect = left_list.rectangle()

        print(f"  HWND=0x{hwnd:08X}")
        print(f"  크기: {rect.width()}x{rect.height()}")

        # 컬럼 좌표 추정
        # 스프레드시트 너비 403px를 4개 컬럼으로 나눔
        # 컬럼: 사번(약 100px), 성명(약 100px), 주민번호(약 150px), 나이(약 50px)
        column_tests = [
            {"x": 50,  "label": "컬럼1(사번추정)", "data": "TEST001"},
            {"x": 100, "label": "컬럼2(성명추정)", "data": "김테스트"},
            {"x": 200, "label": "컬럼3(주민번호추정)", "data": "900101-1234567"},
            {"x": 320, "label": "컬럼4(나이추정)", "data": "35"},
        ]

        print("\n[3/3] 각 컬럼별 테스트 (y=30, 첫 번째 빈 행)...")

        results = []
        for idx, col in enumerate(column_tests):
            print(f"\n테스트 {idx+1}/4: {col['label']}")

            success = single_cell_input(
                hwnd,
                col['x'],
                30,  # y=30 (첫 번째 빈 행)
                col['data'],
                capture_func,
                f"attempt17_02_col{idx+1}_{col['label']}.png"
            )

            results.append({
                "column": col['label'],
                "x": col['x'],
                "data": col['data'],
                "success": success
            })

            time.sleep(0.5)  # 다음 테스트 전 대기

        # 최종 스크린샷
        time.sleep(1)
        capture_func("attempt17_03_final.png")

        # 결과 요약
        print("\n" + "=" * 60)
        print("테스트 결과:")
        print("=" * 60)
        for r in results:
            status = "✅" if r['success'] else "❌"
            print(f"  {status} {r['column']} (x={r['x']}): \"{r['data']}\"")

        print("\n⚠️  스크린샷을 확인하여 각 데이터가 올바른 컬럼에 입력되었는지 확인하세요!")

        return {
            "success": True,
            "message": f"4개 컬럼 테스트 완료",
            "results": results
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
