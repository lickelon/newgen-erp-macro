"""
Attempt 49: pywinauto의 click() 메서드

pywinauto wrapper로 버튼 찾아서 click() 호출
"""
import time
import win32gui


def find_tab_buttons_as_wrappers(dlg):
    """pywinauto wrapper로 탭 버튼들 찾기"""
    buttons = []
    pages = []

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
        return None, None

    # 버튼과 페이지 찾기
    for ctrl in tab_control.children():
        try:
            class_name = ctrl.class_name()
            if class_name == "Button":
                buttons.append(ctrl)
            elif class_name == "#32770":
                text = ctrl.window_text().strip()
                pages.append((ctrl, text))
        except:
            pass

    return buttons, pages


def test_pywinauto_click(dlg, capture_func):
    """pywinauto click() 메서드"""
    print("\n" + "=" * 60)
    print("Attempt 49: pywinauto click()")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/6] 현재 상태...")
        capture_func("attempt49_00_initial.png")
        time.sleep(0.5)

        # 버튼과 페이지 찾기
        buttons, pages = find_tab_buttons_as_wrappers(dlg)

        if not buttons or not pages:
            return {"success": False, "message": "버튼/페이지 없음"}

        print(f"  버튼: {len(buttons)}개")
        print(f"  페이지: {len(pages)}개")
        for idx, (page, text) in enumerate(pages):
            print(f"    [{idx}] {text}")

        # 기본사항 탭 (button[0])
        print("\n[2/6] 기본사항 탭 - pywinauto click()...")
        try:
            buttons[0].click()
            print("  ✓ click() 호출")
            time.sleep(1.0)
            capture_func("attempt49_01_basic_tab.png")
        except Exception as e:
            print(f"  ✗ 오류: {e}")

        # 부양가족명세 탭 (button[1])
        print("\n[3/6] 부양가족명세 탭 - pywinauto click()...")
        try:
            buttons[1].click()
            print("  ✓ click() 호출")
            time.sleep(1.0)
            capture_func("attempt49_02_dependents_tab.png")
        except Exception as e:
            print(f"  ✗ 오류: {e}")

        # 추가사항 탭 (button[2])
        print("\n[4/6] 추가사항 탭 - pywinauto click()...")
        try:
            buttons[2].click()
            print("  ✓ click() 호출")
            time.sleep(1.0)
            capture_func("attempt49_03_extra_tab.png")
        except Exception as e:
            print(f"  ✗ 오류: {e}")

        # 다시 부양가족명세
        print("\n[5/6] 다시 부양가족명세 탭...")
        try:
            buttons[1].click()
            print("  ✓ click() 호출")
            time.sleep(1.0)
            capture_func("attempt49_04_dependents_again.png")
        except Exception as e:
            print(f"  ✗ 오류: {e}")

        # 다시 기본사항
        print("\n[6/6] 다시 기본사항 탭...")
        try:
            buttons[0].click()
            print("  ✓ click() 호출")
            time.sleep(1.0)
            capture_func("attempt49_05_basic_again.png")
        except Exception as e:
            print(f"  ✗ 오류: {e}")

        capture_func("attempt49_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "pywinauto click() 완료"
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

    result = test_pywinauto_click(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
