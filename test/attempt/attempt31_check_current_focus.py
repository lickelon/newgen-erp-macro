"""
Attempt 31: 현재 포커스된 스프레드와 셀 위치 확인

GetFocus()로 현재 포커스된 컨트롤 확인
"""
import time
import win32gui
import win32api
import win32con


def check_current_focus(dlg, capture_func):
    """현재 포커스된 컨트롤 확인"""
    print("\n" + "=" * 60)
    print("Attempt 31: 현재 포커스 확인")
    print("=" * 60)

    try:
        capture_func("attempt31_00_initial.png")

        # 부양가족정보 탭 선택
        print("\n[1/4] 부양가족정보 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("부양가족정보")
        time.sleep(0.5)
        capture_func("attempt31_01_dependents_tab.png")

        # 모든 fpUSpread80 찾기
        print("\n[2/4] 모든 스프레드 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        print(f"  총 {len(spread_controls)}개 fpUSpread80 발견:")
        for idx, ctrl in enumerate(spread_controls):
            rect = ctrl.rectangle()
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            print(f"    Spread #{idx}: HWND=0x{ctrl.handle:08X}, {width}x{height}")

        # 현재 포커스된 윈도우 확인
        print("\n[3/4] 현재 포커스 확인...")
        focused_hwnd = win32gui.GetFocus()
        print(f"  포커스된 HWND: 0x{focused_hwnd:08X}")

        # 포커스된 컨트롤이 어떤 스프레드인지 확인
        focused_spread_idx = None
        for idx, ctrl in enumerate(spread_controls):
            if ctrl.handle == focused_hwnd:
                focused_spread_idx = idx
                print(f"  → Spread #{idx}에 포커스됨!")
                break

        if focused_spread_idx is None:
            # 포커스된 컨트롤의 클래스명 확인
            try:
                class_name = win32gui.GetClassName(focused_hwnd)
                print(f"  → 다른 컨트롤에 포커스: {class_name}")
            except:
                print(f"  → 알 수 없는 컨트롤")

        # 왼쪽 스프레드 클릭 후 포커스 확인
        print("\n[4/4] 왼쪽 스프레드 클릭 후 포커스 재확인...")

        if len(spread_controls) >= 1:
            left_spread = spread_controls[0]  # 첫 번째 = 왼쪽
            hwnd = left_spread.handle

            # 클릭
            lparam = win32api.MAKELONG(100, 100)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.03)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.3)

            capture_func("attempt31_02_after_left_click.png")

            # 포커스 재확인
            focused_hwnd_after = win32gui.GetFocus()
            print(f"  왼쪽 클릭 후 포커스: 0x{focused_hwnd_after:08X}")

            if focused_hwnd_after == hwnd:
                print(f"  → ✅ 왼쪽 스프레드에 포커스!")
            else:
                for idx, ctrl in enumerate(spread_controls):
                    if ctrl.handle == focused_hwnd_after:
                        print(f"  → Spread #{idx}에 포커스")
                        break
                else:
                    try:
                        class_name = win32gui.GetClassName(focused_hwnd_after)
                        print(f"  → 다른 컨트롤: {class_name}")
                    except:
                        print(f"  → 알 수 없음")

        # 오른쪽 스프레드 클릭 후 포커스 확인
        print("\n  오른쪽 영역 클릭 후 포커스 확인...")

        if len(spread_controls) >= 2:
            # 두 번째 또는 세 번째 스프레드 (오른쪽)
            right_spread = spread_controls[1]
            hwnd = right_spread.handle

            # 클릭
            lparam = win32api.MAKELONG(100, 100)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.03)
            win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.3)

            capture_func("attempt31_03_after_right_click.png")

            # 포커스 재확인
            focused_hwnd_after = win32gui.GetFocus()
            print(f"  오른쪽 클릭 후 포커스: 0x{focused_hwnd_after:08X}")

            if focused_hwnd_after == hwnd:
                print(f"  → ✅ 오른쪽 스프레드에 포커스!")
            else:
                for idx, ctrl in enumerate(spread_controls):
                    if ctrl.handle == focused_hwnd_after:
                        print(f"  → Spread #{idx}에 포커스")
                        break

        capture_func("attempt31_04_final.png")

        print("\n" + "=" * 60)
        print("포커스 확인 완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "포커스 확인 완료",
            "spread_count": len(spread_controls),
            "initial_focus": f"0x{focused_hwnd:08X}",
            "focused_spread": focused_spread_idx
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

    result = check_current_focus(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
