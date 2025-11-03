"""
Attempt 11: 화면 좌표로 직접 클릭 + WM_CHAR로 입력

마우스를 움직이지 않고 특정 화면 좌표에 클릭 메시지 전송 후
WM_CHAR로 한 글자씩 입력
"""
import time
import win32api
import win32con
import win32gui


def screen_to_client(hwnd, screen_x, screen_y):
    """화면 좌표를 클라이언트 좌표로 변환"""
    point = win32api.POINT()
    point.x = screen_x
    point.y = screen_y
    win32gui.ScreenToClient(hwnd, point)
    return point.x, point.y


def send_text_via_wmchar(hwnd, text):
    """WM_CHAR로 텍스트 입력"""
    for char in text:
        char_code = ord(char)
        win32api.SendMessage(hwnd, win32con.WM_CHAR, char_code, 0)
        time.sleep(0.02)


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 11: 화면 좌표 클릭 + WM_CHAR 입력")
    print("=" * 60)

    try:
        capture_func("attempt11_00_initial.png")

        # 기본사항 탭
        print("\n[1/5] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt11_01_basic_tab.png")

        main_hwnd = dlg.handle

        # 화면에서 입력 필드의 대략적인 위치 (스크린샷 기준)
        # 실제 좌표는 프로그램 위치에 따라 다름
        print("\n[2/5] 메인 윈도우 위치 확인...")
        main_rect = dlg.rectangle()
        print(f"  메인 윈도우: L={main_rect.left} T={main_rect.top}")
        print(f"              R={main_rect.right} B={main_rect.bottom}")

        # 입력 필드 추정 위치 (상대 좌표)
        # 기본사항 탭 내부의 Edit 필드들
        fields = [
            {"name": "필드1", "offset_x": 200, "offset_y": 300},
            {"name": "필드2", "offset_x": 200, "offset_y": 350},
            {"name": "필드3", "offset_x": 200, "offset_y": 400},
        ]

        test_data = {
            "필드1": "ABC123",
            "필드2": "900101",
            "필드3": "홍길동"
        }

        print(f"\n[3/5] 테스트 데이터: {test_data}")

        print("\n[4/5] 좌표 기반 입력 시도...")

        # 먼저 WindowFromPoint로 실제 컨트롤 찾기
        print("\n  실제 컨트롤 좌표 스캔...")
        scan_points = [
            (main_rect.left + 200, main_rect.top + 250),
            (main_rect.left + 200, main_rect.top + 300),
            (main_rect.left + 200, main_rect.top + 350),
            (main_rect.left + 300, main_rect.top + 300),
            (main_rect.left + 400, main_rect.top + 300),
        ]

        found_controls = {}
        for screen_x, screen_y in scan_points:
            try:
                ctrl_hwnd = win32gui.WindowFromPoint((screen_x, screen_y))
                if ctrl_hwnd and ctrl_hwnd != main_hwnd:
                    class_name = win32gui.GetClassName(ctrl_hwnd)
                    coord_key = f"({screen_x},{screen_y})"
                    found_controls[coord_key] = {
                        "hwnd": ctrl_hwnd,
                        "class": class_name
                    }
                    print(f"    {coord_key} → HWND=0x{ctrl_hwnd:08X} class={class_name}")
            except:
                pass

        # Edit 컨트롤 찾기
        edit_hwnds = []
        for coord, info in found_controls.items():
            if "Edit" in info["class"]:
                edit_hwnds.append(info["hwnd"])

        print(f"\n  Edit 컨트롤 {len(set(edit_hwnds))}개 발견")
        unique_edits = list(set(edit_hwnds))

        if not unique_edits:
            print("  ⚠️  Edit 컨트롤을 찾을 수 없습니다")
            print("  대신 메인 윈도우에 키보드 입력 시도...")

            # 메인 윈도우에 포커스 설정 시도
            try:
                win32gui.SetForegroundWindow(main_hwnd)
                time.sleep(0.2)
            except:
                pass

            # Tab 키로 필드 이동 + 입력
            print("\n  Tab 키로 필드 탐색...")
            for field_name, value in test_data.items():
                print(f"    {field_name}: \"{value}\"")

                # Tab 키 (다음 필드로)
                win32api.SendMessage(main_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                time.sleep(0.05)
                win32api.SendMessage(main_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                time.sleep(0.1)

                # 텍스트 입력
                send_text_via_wmchar(main_hwnd, value)
                time.sleep(0.2)

        else:
            # Edit 컨트롤에 직접 입력
            print("\n  발견된 Edit 컨트롤에 입력...")
            for idx, hwnd in enumerate(unique_edits):
                if idx >= len(test_data):
                    break

                field_name = list(test_data.keys())[idx]
                value = test_data[field_name]

                print(f"    [{idx}] {field_name}: \"{value}\" → HWND=0x{hwnd:08X}")

                try:
                    # WM_SETTEXT 시도
                    win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, value)
                    time.sleep(0.1)

                    # 확인
                    text = win32gui.GetWindowText(hwnd)
                    print(f"         결과: \"{text}\" {'✅' if text == value else '❌'}")

                except Exception as e:
                    print(f"         ✗ 오류: {e}")

        time.sleep(1)
        capture_func("attempt11_02_after_input.png")

        return {"success": True, "message": "좌표 기반 입력 시도 완료"}

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
