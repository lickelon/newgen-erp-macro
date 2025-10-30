"""
Attempt 44: 특정 윈도우 핸들 클릭 테스트

Window 0x000808EE와 0x000809D4 클릭
"""
import time
import win32api
import win32con
import win32gui


def test_click_handles(dlg, capture_func):
    """특정 핸들 클릭 테스트"""
    print("\n" + "=" * 60)
    print("Attempt 44: 특정 핸들 클릭")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/5] 현재 상태...")
        capture_func("attempt44_00_initial.png")
        time.sleep(0.5)

        # 핸들 목록
        handles = [0x000808EE, 0x000809D4]

        for idx, hwnd in enumerate(handles):
            print(f"\n[{idx+2}/5] HWND 0x{hwnd:08X} 클릭 시도...")

            # 윈도우 존재 확인
            if not win32gui.IsWindow(hwnd):
                print(f"  ✗ 윈도우가 존재하지 않음")
                continue

            print(f"  ✓ 윈도우 존재 확인")

            # 중앙 좌표 클릭 (10, 10)
            lparam = win32api.MAKELONG(10, 10)

            print(f"  → WM_LBUTTONDOWN 전송...")
            result_down = win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            print(f"     결과: {result_down}")
            time.sleep(0.1)

            print(f"  → WM_LBUTTONUP 전송...")
            result_up = win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            print(f"     결과: {result_up}")
            time.sleep(1.0)

            capture_func(f"attempt44_{idx+1:02d}_after_click_0x{hwnd:08X}.png")

        print("\n[5/5] 최종 상태...")
        time.sleep(0.5)
        capture_func("attempt44_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "핸들 클릭 테스트 완료"
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

    result = test_click_handles(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
