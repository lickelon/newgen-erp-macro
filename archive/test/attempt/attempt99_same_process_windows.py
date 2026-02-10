"""
시도 99: 같은 프로세스의 모든 창 찾기
"""
from pywinauto import Application
import win32process
import win32gui

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 99: 같은 프로세스의 모든 창 찾기")
    print("="*60)

    try:
        # 급여자료입력 연결
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        # 프로세스 ID 가져오기
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)
        print(f"\n급여자료입력")
        print(f"  HWND: 0x{main_window.handle:08X}")
        print(f"  PID: {process_id}")

        # 같은 프로세스의 모든 창 찾기
        print(f"\n같은 프로세스(PID={process_id})의 모든 창:")
        print("-" * 60)

        found_windows = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        class_name = win32gui.GetClassName(hwnd)
                        results.append((hwnd, title, class_name))
                    except:
                        pass
            return True

        win32gui.EnumWindows(enum_callback, found_windows)

        print(f"총 {len(found_windows)}개 창 발견:")

        found_installment = False
        for hwnd, title, class_name in found_windows:
            print(f"\n  HWND: 0x{hwnd:08X}")
            print(f"  클래스: {class_name}")
            if title:
                print(f"  제목: '{title}'")
                if "분납" in title:
                    print(f"  >>> ⭐⭐⭐ 분납 관련 창 발견!")
                    found_installment = True

        if found_installment:
            return {"success": True, "message": "분납 관련 창 발견"}
        else:
            return {"success": False, "message": "분납 관련 창을 찾지 못했습니다"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
