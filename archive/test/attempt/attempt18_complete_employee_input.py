"""
Attempt 18: 완전한 직원 정보 입력 통합 테스트

확립된 방법으로 사번, 성명, 주민번호, 나이를 순차적으로 입력
각 단계마다 스크린샷으로 확인
"""
import time
import win32api
import win32con


def input_to_spread(hwnd, x, y, text, label, capture_func, screenshot_name):
    """fpUSpread에 입력 + 스크린샷"""
    print(f"  {label}: \"{text}\" at ({x},{y})")

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
        time.sleep(0.5)

        # 스크린샷
        capture_func(screenshot_name)
        print(f"    ✓ 입력 완료, 스크린샷: {screenshot_name}")

        return True

    except Exception as e:
        print(f"    ✗ 오류: {e}")
        return False


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 18: 완전한 직원 정보 입력")
    print("=" * 60)

    try:
        capture_func("attempt18_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt18_01_basic_tab.png")

        # fpUSpread80 찾기
        print("\n[2/4] fpUSpread 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            return {"success": False, "message": "fpUSpread 부족"}

        # Spread #2 (왼쪽 목록) 사용
        left_list = spread_controls[2]
        hwnd = left_list.handle

        print(f"  Spread #2 (왼쪽 목록): HWND=0x{hwnd:08X}")

        # 완전한 직원 정보
        employee_data = {
            "사번": "2025100",
            "성명": "테스트사원",
            "주민번호": "920315-1234567",
            "나이": "33"
        }

        print(f"\n[3/4] 직원 정보 입력:")
        print(f"  {employee_data}")

        # 컬럼 좌표 (Attempt 17에서 확인된 값)
        # y=30은 새 행 추가, y값을 다르게 해서 각 필드를 별도 행에 입력
        # 또는 x값만 바꿔서 같은 행의 다른 컬럼에 입력

        # 방법 1: 같은 행(y=30)의 다른 컬럼들에 순차 입력
        inputs = [
            (50,  30, employee_data["사번"], "사번", "attempt18_02_sabun.png"),
            (100, 30, employee_data["성명"], "성명", "attempt18_03_name.png"),
            (200, 30, employee_data["주민번호"], "주민번호", "attempt18_04_jumin.png"),
            (320, 30, employee_data["나이"], "나이", "attempt18_05_age.png"),
        ]

        success_count = 0
        for x, y, value, label, screenshot in inputs:
            if input_to_spread(hwnd, x, y, value, label, capture_func, screenshot):
                success_count += 1
            time.sleep(0.3)  # 각 입력 사이 대기

        # 최종 상태
        time.sleep(1)
        capture_func("attempt18_06_final.png")

        print(f"\n[4/4] 결과: {success_count}/4개 필드 입력 완료")
        print("  ⚠️  스크린샷을 확인하여 데이터 입력 결과를 확인하세요!")

        return {
            "success": success_count >= 3,
            "message": f"{success_count}/4개 필드 입력",
            "employee_data": employee_data
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
