"""
Attempt 46: 탭 컨트롤의 자식 윈도우 구조 출력

탭 버튼들이 어디에 있는지 확인
"""
import win32gui


def print_window_tree(hwnd, indent=0):
    """윈도우 트리 출력"""
    try:
        class_name = win32gui.GetClassName(hwnd)
        text = win32gui.GetWindowText(hwnd)
        print(f"{'  ' * indent}0x{hwnd:08X} [{class_name}] \"{text}\"")
    except:
        print(f"{'  ' * indent}0x{hwnd:08X} [ERROR]")

    # 자식 윈도우 열거
    def enum_child(child_hwnd, lparam):
        print_window_tree(child_hwnd, indent + 1)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, enum_child, None)
    except:
        pass


def test_print_tab_children(dlg, capture_func):
    """탭 컨트롤 자식 출력"""
    print("\n" + "=" * 60)
    print("Attempt 46: 탭 컨트롤 자식 윈도우 출력")
    print("=" * 60)

    try:
        # 탭 컨트롤 찾기
        print("\n탭 컨트롤 찾기...")
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

        hwnd = tab_control.handle
        print(f"✓ 탭 컨트롤: 0x{hwnd:08X}")
        print(f"  클래스: {tab_control.class_name()}")
        print(f"  텍스트: {tab_control.window_text()}")

        print("\n" + "=" * 60)
        print("윈도우 트리:")
        print("=" * 60)
        print_window_tree(hwnd)

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "출력 완료"
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

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    result = test_print_tab_children(dlg, None)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
