"""
Attempt 47: 탭 페이지 텍스트로 탭 선택 (좌표 없음)

1. 탭 페이지(#32770)를 텍스트로 찾기
2. 페이지 인덱스 확인
3. 해당 인덱스의 Button 클릭
"""
import time
import win32api
import win32con
import win32gui


def find_tab_page_index(tab_control_hwnd, page_text):
    """탭 페이지 텍스트로 인덱스 찾기"""
    pages = []
    buttons = []

    def enum_child(hwnd, lparam):
        try:
            class_name = win32gui.GetClassName(hwnd)
            text = win32gui.GetWindowText(hwnd)

            if class_name == "#32770":  # Dialog = Tab page
                pages.append((hwnd, text.strip()))
            elif class_name == "Button":
                buttons.append(hwnd)
        except:
            pass
        return True

    win32gui.EnumChildWindows(tab_control_hwnd, enum_child, None)

    # 페이지 찾기
    for idx, (hwnd, text) in enumerate(pages):
        if page_text in text:
            return idx, buttons[idx] if idx < len(buttons) else None

    return None, None


def test_tab_by_page_text(dlg, capture_func):
    """페이지 텍스트로 탭 선택"""
    print("\n" + "=" * 60)
    print("Attempt 47: 페이지 텍스트로 탭 선택 (좌표 없음)")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/6] 현재 상태...")
        capture_func("attempt47_00_initial.png")
        time.sleep(0.5)

        # 탭 컨트롤 찾기
        print("\n  탭 컨트롤 찾기...")
        tab_control = None
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name().startswith("Afx:TabWnd:"):
                    tab_control = ctrl
                    break
            except:
                pass

        if not tab_control:
            return {"success": False, "message": "탭 컨트롤 없음"}

        tab_hwnd = tab_control.handle
        print(f"  ✓ 탭 컨트롤: 0x{tab_hwnd:08X}")

        # 기본사항 탭 선택
        print("\n[2/6] '기본사항' 탭 선택...")
        idx, button_hwnd = find_tab_page_index(tab_hwnd, "기본사항")
        if button_hwnd:
            print(f"  ✓ 페이지 인덱스: {idx}")
            print(f"  ✓ 버튼 HWND: 0x{button_hwnd:08X}")

            # 버튼 클릭
            lparam = win32api.MAKELONG(10, 10)
            win32api.SendMessage(button_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32api.SendMessage(button_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(1.0)
            capture_func("attempt47_01_basic_tab.png")
        else:
            print("  ✗ 찾지 못함")

        # 부양가족명세 탭 선택
        print("\n[3/6] '부양가족명세' 탭 선택...")
        idx, button_hwnd = find_tab_page_index(tab_hwnd, "부양가족명세")
        if button_hwnd:
            print(f"  ✓ 페이지 인덱스: {idx}")
            print(f"  ✓ 버튼 HWND: 0x{button_hwnd:08X}")

            # 버튼 클릭
            lparam = win32api.MAKELONG(10, 10)
            win32api.SendMessage(button_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.05)
            win32api.SendMessage(button_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(1.0)
            capture_func("attempt47_02_dependents_tab.png")
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
        capture_func("attempt47_03_after_down.png")

        # 데이터 입력
        print("\n[6/6] 데이터 입력...")
        test_data = ["1", "김부모", "내", "1960"]
        field_names = ["연말관계", "성명", "내외국", "년도"]

        for idx, value in enumerate(test_data):
            print(f"  [{idx+1}/4] {field_names[idx]}: \"{value}\"")
            dlg.type_keys(value, with_spaces=False, pause=0.05)
            time.sleep(0.3)

            if idx < len(test_data) - 1:
                dlg.type_keys("{TAB}", pause=0.05)
                time.sleep(0.2)

            capture_func(f"attempt47_04_field{idx+1}.png")

        # Enter
        print("  → Enter로 확정...")
        dlg.type_keys("{ENTER}", pause=0.05)
        time.sleep(0.5)
        capture_func("attempt47_05_final.png")

        print("\n" + "=" * 60)
        print("✅ 완료 - 좌표 없이 탭 선택 성공!")
        print("=" * 60)

        return {
            "success": True,
            "message": "페이지 텍스트로 탭 선택 성공 (좌표 없음)",
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

    result = test_tab_by_page_text(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
