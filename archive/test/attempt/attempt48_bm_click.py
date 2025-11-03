"""
Attempt 48: BM_CLICK 메시지로 버튼 클릭

WM_LBUTTONDOWN/UP 대신 BM_CLICK 메시지 사용
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

            if class_name == "#32770":
                pages.append((hwnd, text.strip()))
            elif class_name == "Button":
                buttons.append(hwnd)
        except:
            pass
        return True

    win32gui.EnumChildWindows(tab_control_hwnd, enum_child, None)

    for idx, (hwnd, text) in enumerate(pages):
        if page_text in text:
            return idx, buttons[idx] if idx < len(buttons) else None

    return None, None


def test_bm_click(dlg, capture_func):
    """BM_CLICK 메시지로 탭 클릭"""
    print("\n" + "=" * 60)
    print("Attempt 48: BM_CLICK 메시지")
    print("=" * 60)

    BM_CLICK = 0x00F5

    try:
        # 현재 상태
        print("\n[1/5] 현재 상태...")
        capture_func("attempt48_00_initial.png")
        time.sleep(0.5)

        # 탭 컨트롤 찾기
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
        print(f"  탭 컨트롤: 0x{tab_hwnd:08X}")

        # 기본사항 탭 선택 (BM_CLICK)
        print("\n[2/5] '기본사항' 탭 - BM_CLICK...")
        idx, button_hwnd = find_tab_page_index(tab_hwnd, "기본사항")
        if button_hwnd:
            print(f"  버튼: 0x{button_hwnd:08X}")
            result = win32api.SendMessage(button_hwnd, BM_CLICK, 0, 0)
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt48_01_basic_tab.png")
        else:
            print("  ✗ 찾지 못함")

        # 부양가족명세 탭 선택 (BM_CLICK)
        print("\n[3/5] '부양가족명세' 탭 - BM_CLICK...")
        idx, button_hwnd = find_tab_page_index(tab_hwnd, "부양가족명세")
        if button_hwnd:
            print(f"  버튼: 0x{button_hwnd:08X}")
            result = win32api.SendMessage(button_hwnd, BM_CLICK, 0, 0)
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt48_02_dependents_tab.png")
        else:
            print("  ✗ 찾지 못함")

        # 추가사항 탭 선택 (BM_CLICK)
        print("\n[4/5] '추가사항' 탭 - BM_CLICK...")
        idx, button_hwnd = find_tab_page_index(tab_hwnd, "추가사항")
        if button_hwnd:
            print(f"  버튼: 0x{button_hwnd:08X}")
            result = win32api.SendMessage(button_hwnd, BM_CLICK, 0, 0)
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt48_03_extra_tab.png")
        else:
            print("  ✗ 찾지 못함")

        # 다시 기본사항
        print("\n[5/5] 다시 '기본사항' 탭...")
        idx, button_hwnd = find_tab_page_index(tab_hwnd, "기본사항")
        if button_hwnd:
            print(f"  버튼: 0x{button_hwnd:08X}")
            result = win32api.SendMessage(button_hwnd, BM_CLICK, 0, 0)
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt48_04_basic_again.png")

        capture_func("attempt48_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "BM_CLICK 방식 완료"
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

    result = test_bm_click(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
