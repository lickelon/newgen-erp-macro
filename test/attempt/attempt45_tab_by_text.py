"""
Attempt 45: 탭 텍스트로 버튼 찾아서 클릭 (좌표 없음)

Window 트리에서 Button 컨트롤의 텍스트를 확인하여 탭 선택
"""
import time
import win32api
import win32con
import win32gui


def find_tab_button_by_text(parent_hwnd, tab_text):
    """텍스트로 탭 버튼 찾기"""
    found_hwnd = None

    def enum_child_proc(hwnd, lparam):
        nonlocal found_hwnd
        try:
            # 윈도우 텍스트 가져오기
            length = win32gui.GetWindowTextLength(hwnd)
            if length > 0:
                text = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)

                # Button 클래스이고 텍스트가 일치하면
                if class_name == "Button" and tab_text in text:
                    found_hwnd = hwnd
                    return False  # 찾았으면 중단
        except:
            pass
        return True

    win32gui.EnumChildWindows(parent_hwnd, enum_child_proc, None)
    return found_hwnd


def test_tab_by_text(dlg, capture_func):
    """텍스트로 탭 선택 테스트"""
    print("\n" + "=" * 60)
    print("Attempt 45: 텍스트로 탭 선택 (좌표 없음)")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/6] 현재 상태...")
        capture_func("attempt45_00_initial.png")
        time.sleep(0.5)

        dlg_hwnd = dlg.handle
        print(f"  Dialog HWND: 0x{dlg_hwnd:08X}")

        # 탭 컨트롤 찾기
        print("\n  탭 컨트롤 찾기...")
        tab_control_hwnd = None
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name().startswith("Afx:TabWnd:"):
                    tab_control_hwnd = ctrl.handle
                    print(f"  ✓ 탭 컨트롤: 0x{tab_control_hwnd:08X}")
                    break
            except:
                pass

        if not tab_control_hwnd:
            return {"success": False, "message": "탭 컨트롤 없음"}

        # 기본사항 탭 찾기
        print("\n[2/6] '기본사항' 탭 버튼 찾기...")
        basic_tab = find_tab_button_by_text(tab_control_hwnd, "기본사항")
        if basic_tab:
            print(f"  ✓ 찾음: 0x{basic_tab:08X}")
            print(f"     텍스트: {win32gui.GetWindowText(basic_tab)}")

            # 클릭
            print("  → 클릭...")
            lparam = win32api.MAKELONG(10, 10)
            win32api.SendMessage(basic_tab, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32api.SendMessage(basic_tab, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(1.0)
            capture_func("attempt45_01_basic_tab.png")
        else:
            print("  ✗ 찾지 못함")

        # 부양가족정보 탭 찾기
        print("\n[3/6] '부양가족정보' 탭 버튼 찾기...")
        dependents_tab = find_tab_button_by_text(tab_control_hwnd, "부양가족정보")
        if dependents_tab:
            print(f"  ✓ 찾음: 0x{dependents_tab:08X}")
            print(f"     텍스트: {win32gui.GetWindowText(dependents_tab)}")

            # 클릭
            print("  → 클릭...")
            lparam = win32api.MAKELONG(10, 10)
            win32api.SendMessage(dependents_tab, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32api.SendMessage(dependents_tab, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(1.0)
            capture_func("attempt45_02_dependents_tab.png")
        else:
            print("  ✗ 찾지 못함")

        # 스프레드 찾기
        print("\n[4/6] 스프레드 찾기...")
        spread = None
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread = ctrl
                    break
            except:
                pass

        if not spread:
            return {"success": False, "message": "스프레드 없음"}

        hwnd = spread.handle
        print(f"  Spread HWND: 0x{hwnd:08X}")

        # Down 키
        print("\n[5/6] Down 키로 행 이동...")
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        time.sleep(0.02)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.5)
        capture_func("attempt45_03_after_down.png")

        # 데이터 입력
        print("\n[6/6] 데이터 입력...")
        test_data = ["3", "김배우자", "내", "1988"]
        field_names = ["연말관계", "성명", "내외국", "년도"]

        for idx, value in enumerate(test_data):
            print(f"  [{idx+1}/4] {field_names[idx]}: \"{value}\"")
            dlg.type_keys(value, with_spaces=False, pause=0.05)
            time.sleep(0.3)

            if idx < len(test_data) - 1:
                dlg.type_keys("{TAB}", pause=0.05)
                time.sleep(0.2)

            capture_func(f"attempt45_04_field{idx+1}.png")

        # Enter
        print("  → Enter로 확정...")
        dlg.type_keys("{ENTER}", pause=0.05)
        time.sleep(0.5)
        capture_func("attempt45_05_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "텍스트 기반 탭 선택 완료 (좌표 없음)",
            "data": test_data
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

    result = test_tab_by_text(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
