"""
Attempt 13: fpUSpread80 여러 셀 좌표 시도

성공한 방법(Attempt 12)을 사용하여 여러 셀에 입력 시도
다양한 좌표를 테스트하여 사번, 성명, 주민번호 필드 찾기
"""
import time
import win32api
import win32con


def click_and_input(hwnd, x, y, text, label):
    """fpUSpread80 셀 클릭 + 텍스트 입력"""
    print(f"  [{label}] 좌표=({x},{y}) 텍스트=\"{text}\"")

    try:
        # 클릭 (셀 선택)
        lparam = win32api.MAKELONG(x, y)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.05)
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.2)

        # 텍스트 입력 (WM_CHAR)
        for char in text:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.02)

        # Enter
        time.sleep(0.1)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.3)

        print(f"       ✓ 입력 완료")
        return True

    except Exception as e:
        print(f"       ✗ 오류: {e}")
        return False


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 13: fpUSpread80 여러 셀에 입력")
    print("=" * 60)

    try:
        capture_func("attempt13_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt13_01_basic_tab.png")

        # fpUSpread80 찾기 (기본사항 탭의 것)
        print("\n[2/4] fpUSpread80 컨트롤 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            return {"success": False, "message": "fpUSpread80 부족"}

        # 세 번째 컨트롤 (기본사항 탭의 메인 스프레드)
        spread_ctrl = spread_controls[2]
        spread_hwnd = spread_ctrl.handle
        rect = spread_ctrl.rectangle()

        print(f"  대상: HWND=0x{spread_hwnd:08X}")
        print(f"       위치: ({rect.left},{rect.top})")
        print(f"       크기: {rect.width()}x{rect.height()}")

        # 테스트할 좌표들 (fpUSpread 내부 클라이언트 좌표)
        # 스프레드시트는 보통 행/열 구조
        test_cells = [
            # 상단 영역
            {"x": 50, "y": 30, "text": "사번001", "label": "좌상단"},
            {"x": 150, "y": 30, "text": "김철수", "label": "상단중앙"},
            {"x": 250, "y": 30, "text": "900101", "label": "상단우측"},

            # 중간 영역
            {"x": 50, "y": 80, "text": "사번002", "label": "중상단"},
            {"x": 150, "y": 80, "text": "이영희", "label": "중상중앙"},
            {"x": 250, "y": 80, "text": "900201", "label": "중상우측"},

            # 하단 영역
            {"x": 50, "y": 150, "text": "사번003", "label": "중하단"},
            {"x": 150, "y": 150, "text": "박민수", "label": "중하중앙"},
            {"x": 250, "y": 150, "text": "900301", "label": "중하우측"},

            # 더 아래
            {"x": 50, "y": 250, "text": "사번004", "label": "하단"},
        ]

        print(f"\n[3/4] 여러 셀에 입력 시도 ({len(test_cells)}개)...")

        success_count = 0
        for cell in test_cells:
            if click_and_input(spread_hwnd, cell["x"], cell["y"], cell["text"], cell["label"]):
                success_count += 1

        time.sleep(1)
        capture_func("attempt13_02_after_input.png")

        print(f"\n[4/4] 결과: {success_count}/{len(test_cells)}개 입력 완료")
        print("  ⚠️  스크린샷에서 실제 어디에 입력됐는지 확인 필요")

        return {
            "success": success_count > 0,
            "message": f"{success_count}/{len(test_cells)}개 셀 입력 완료"
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
