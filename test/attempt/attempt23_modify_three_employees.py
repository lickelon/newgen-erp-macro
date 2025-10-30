"""
Attempt 23: 기존 직원 3명 수정 (키보드만 사용)

좌표 없이 키보드 네비게이션으로 1, 2, 3번 직원 정보 수정
"""
import time
import win32api
import win32con


def modify_employees(dlg, capture_func):
    """왼쪽 스프레드에서 3명의 직원 정보 수정"""
    print("\n" + "=" * 60)
    print("Attempt 23: 기존 직원 3명 수정")
    print("=" * 60)

    try:
        capture_func("attempt23_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt23_01_basic_tab.png")

        # 왼쪽 스프레드 찾기
        print("\n[2/4] 왼쪽 스프레드 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            return {"success": False, "message": "Spread 부족"}

        left_list = spread_controls[2]
        hwnd = left_list.handle
        print(f"  Spread #2 HWND: 0x{hwnd:08X}")

        # 포커스 및 시작 위치로 이동
        print("\n[3/4] 스프레드 포커스 및 시작 위치 이동...")

        # 중앙 클릭으로 포커스
        rect = left_list.rectangle()
        center_x = (rect.right - rect.left) // 2
        center_y = (rect.bottom - rect.top) // 2
        lparam = win32api.MAKELONG(center_x, center_y)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.03)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.3)

        # Ctrl+Home으로 A1
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        # Down으로 첫 데이터 행 (234번 직원)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.3)

        capture_func("attempt23_02_first_row.png")

        # 3명의 직원 데이터
        employees = [
            {
                "name": "첫번째직원",
                "data": ["EMP001", "김첫번", "900101-1111111", "35"]
            },
            {
                "name": "두번째직원",
                "data": ["EMP002", "이두번", "910202-2222222", "34"]
            },
            {
                "name": "세번째직원",
                "data": ["EMP003", "박세번", "920303-1333333", "33"]
            }
        ]

        print("\n[4/4] 3명의 직원 정보 수정...")
        modified_count = 0

        for emp_idx, employee in enumerate(employees):
            emp_name = employee["name"]
            emp_data = employee["data"]

            print(f"\n  [{emp_idx+1}/3] {emp_name} 수정:")

            # 각 필드 입력
            for field_idx, value in enumerate(emp_data):
                field_names = ["사번", "성명", "주민번호", "나이"]
                field_name = field_names[field_idx]

                print(f"    {field_name}: \"{value}\"")

                # 입력
                for char in value:
                    win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                    time.sleep(0.015)

                time.sleep(0.15)

                # Tab으로 다음 필드 (마지막 필드가 아니면)
                if field_idx < len(emp_data) - 1:
                    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                    time.sleep(0.02)
                    win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                    time.sleep(0.15)

            # 스크린샷
            capture_func(f"attempt23_03_employee{emp_idx+1}.png")

            # Enter로 다음 행 (마지막 직원이 아니면)
            if emp_idx < len(employees) - 1:
                print(f"    → Enter (다음 행)")
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                time.sleep(0.3)
            else:
                # 마지막 직원은 Enter로 확정
                print(f"    → Enter (확정)")
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                time.sleep(0.5)

            modified_count += 1

        capture_func("attempt23_04_final.png")

        print("\n" + "=" * 60)
        print(f"완료: {modified_count}명의 직원 정보 수정")
        print("스크린샷 확인 필요")
        print("=" * 60)

        return {
            "success": True,
            "message": f"{modified_count}명 수정 완료",
            "employees": employees
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }


if __name__ == "__main__":
    from pywinauto import application
    import sys
    import os

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    result = modify_employees(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
