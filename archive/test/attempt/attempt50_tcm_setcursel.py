"""
Attempt 50: TCM_SETCURSEL 메시지로 탭 선택

Windows Tab Control 표준 메시지 사용
"""
import time
import win32api
import win32con
import win32gui


def find_tab_page_index(tab_control_hwnd, page_text):
    """탭 페이지 텍스트로 인덱스 찾기"""
    pages = []

    def enum_child(hwnd, lparam):
        try:
            class_name = win32gui.GetClassName(hwnd)
            text = win32gui.GetWindowText(hwnd)

            if class_name == "#32770":
                pages.append((hwnd, text.strip()))
        except:
            pass
        return True

    win32gui.EnumChildWindows(tab_control_hwnd, enum_child, None)

    for idx, (hwnd, text) in enumerate(pages):
        if page_text in text:
            return idx

    return None


def test_tcm_setcursel(dlg, capture_func):
    """TCM_SETCURSEL 메시지"""
    print("\n" + "=" * 60)
    print("Attempt 50: TCM_SETCURSEL 메시지")
    print("=" * 60)

    TCM_SETCURSEL = 0x1300 + 12  # TCM_FIRST + 12
    TCM_GETCURSEL = 0x1300 + 11  # TCM_FIRST + 11

    try:
        # 현재 상태
        print("\n[1/6] 현재 상태...")
        capture_func("attempt50_00_initial.png")
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

        # 현재 선택된 탭 인덱스 확인
        current_idx = win32api.SendMessage(tab_hwnd, TCM_GETCURSEL, 0, 0)
        print(f"  현재 탭 인덱스: {current_idx}")

        # 기본사항 탭 (인덱스 0)
        print("\n[2/6] 기본사항 탭 선택 (인덱스 0)...")
        idx = find_tab_page_index(tab_hwnd, "기본사항")
        if idx is not None:
            print(f"  페이지 인덱스: {idx}")
            result = win32api.SendMessage(tab_hwnd, TCM_SETCURSEL, idx, 0)
            print(f"  TCM_SETCURSEL 결과: {result}")
            time.sleep(1.0)
            capture_func("attempt50_01_basic_tab.png")

            # 확인
            current_idx = win32api.SendMessage(tab_hwnd, TCM_GETCURSEL, 0, 0)
            print(f"  현재 탭 인덱스: {current_idx}")

        # 부양가족명세 탭 (인덱스 1)
        print("\n[3/6] 부양가족명세 탭 선택 (인덱스 1)...")
        idx = find_tab_page_index(tab_hwnd, "부양가족명세")
        if idx is not None:
            print(f"  페이지 인덱스: {idx}")
            result = win32api.SendMessage(tab_hwnd, TCM_SETCURSEL, idx, 0)
            print(f"  TCM_SETCURSEL 결과: {result}")
            time.sleep(1.0)
            capture_func("attempt50_02_dependents_tab.png")

            current_idx = win32api.SendMessage(tab_hwnd, TCM_GETCURSEL, 0, 0)
            print(f"  현재 탭 인덱스: {current_idx}")

        # 추가사항 탭 (인덱스 2)
        print("\n[4/6] 추가사항 탭 선택 (인덱스 2)...")
        idx = find_tab_page_index(tab_hwnd, "추가사항")
        if idx is not None:
            print(f"  페이지 인덱스: {idx}")
            result = win32api.SendMessage(tab_hwnd, TCM_SETCURSEL, idx, 0)
            print(f"  TCM_SETCURSEL 결과: {result}")
            time.sleep(1.0)
            capture_func("attempt50_03_extra_tab.png")

            current_idx = win32api.SendMessage(tab_hwnd, TCM_GETCURSEL, 0, 0)
            print(f"  현재 탭 인덱스: {current_idx}")

        # 다시 부양가족명세
        print("\n[5/6] 다시 부양가족명세...")
        idx = find_tab_page_index(tab_hwnd, "부양가족명세")
        if idx is not None:
            result = win32api.SendMessage(tab_hwnd, TCM_SETCURSEL, idx, 0)
            print(f"  TCM_SETCURSEL 결과: {result}")
            time.sleep(1.0)
            capture_func("attempt50_04_dependents_again.png")

        # 다시 기본사항
        print("\n[6/6] 다시 기본사항...")
        idx = find_tab_page_index(tab_hwnd, "기본사항")
        if idx is not None:
            result = win32api.SendMessage(tab_hwnd, TCM_SETCURSEL, idx, 0)
            print(f"  TCM_SETCURSEL 결과: {result}")
            time.sleep(1.0)
            capture_func("attempt50_05_basic_again.png")

        capture_func("attempt50_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "TCM_SETCURSEL 방식 완료"
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

    result = test_tcm_setcursel(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
