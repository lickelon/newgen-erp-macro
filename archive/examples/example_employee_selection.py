"""
Attempt 53: 사원을 좌표 없이 선택

왼쪽 스프레드시트에서 특정 사원을 키보드로 선택
"""
import time
import win32gui
import win32con
import win32api


def test_select_employee(dlg, capture_func):
    """사원 선택 테스트"""
    print("\n" + "=" * 60)
    print("Attempt 53: 사원 선택 (좌표 없음)")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/10] 현재 상태...")
        capture_func("attempt53_00_initial.png")
        time.sleep(0.5)

        # 모든 스프레드 찾기
        print("\n[2/10] 스프레드 찾기...")
        spreads = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spreads.append(ctrl)
                    rect = ctrl.rectangle()
                    print(f"  스프레드 발견: 0x{ctrl.handle:08X}")
                    print(f"    위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")
                    print(f"    크기: W={rect.width()}, H={rect.height()}")
            except Exception as e:
                pass

        if len(spreads) < 2:
            return {"success": False, "message": "스프레드 부족"}

        # 왼쪽 스프레드 (사원 목록) = 더 작은 X 좌표
        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]  # 사원 목록
        right_spread = spreads[1]  # 부양가족 목록

        print(f"\n  왼쪽 스프레드 (사원): 0x{left_spread.handle:08X}")
        print(f"  오른쪽 스프레드 (부양가족): 0x{right_spread.handle:08X}")

        # 방법 1: 왼쪽 스프레드에 포커스 + Home 키
        print("\n[3/10] 방법 1: 왼쪽 스프레드 set_focus + Home")
        left_spread.set_focus()
        time.sleep(0.5)
        capture_func("attempt53_01_left_focus.png")

        # Home 키
        dlg.type_keys("{HOME}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_02_home.png")

        # 방법 2: Down 키로 이동
        print("\n[4/10] 방법 2: Down 키 2번")
        dlg.type_keys("{DOWN}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_03_down1.png")

        dlg.type_keys("{DOWN}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_04_down2.png")

        # 방법 3: Up 키로 복귀
        print("\n[5/10] 방법 3: Up 키")
        dlg.type_keys("{UP}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_05_up.png")

        # 방법 4: Ctrl+Home (첫 번째 셀)
        print("\n[6/10] 방법 4: Ctrl+Home")
        dlg.type_keys("^{HOME}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_06_ctrl_home.png")

        # 방법 5: SendMessage로 Down
        print("\n[7/10] 방법 5: SendMessage VK_DOWN")
        hwnd = left_spread.handle
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.5)
        capture_func("attempt53_07_sendmsg_down.png")

        # 방법 6: 여러 번 Down
        print("\n[8/10] 방법 6: Down 3번 연속")
        for i in range(3):
            dlg.type_keys("{DOWN}", pause=0.1)
            time.sleep(0.3)
        capture_func("attempt53_08_multi_down.png")

        # 방법 7: PageDown
        print("\n[9/10] 방법 7: PageDown")
        dlg.type_keys("{PGDN}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_09_pagedown.png")

        # 방법 8: Home으로 복귀
        print("\n[10/10] 방법 8: Home 복귀")
        dlg.type_keys("{HOME}", pause=0.1)
        time.sleep(0.5)
        capture_func("attempt53_10_final_home.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "사원 선택 테스트 완료"
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

    result = test_select_employee(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
