"""
Attempt 32: 지금 현재 상태 스크린샷만 찍기

아무것도 하지 않고 현재 상태만 확인
"""
import time


def check_current_state(dlg, capture_func):
    """현재 상태만 확인"""
    print("\n" + "=" * 60)
    print("Attempt 32: 현재 상태 확인")
    print("=" * 60)

    try:
        # 아무것도 하지 않고 스크린샷만
        print("\n현재 상태 스크린샷 촬영...")
        capture_func("attempt32_00_current_state.png")
        time.sleep(0.5)

        print("완료!")

        return {
            "success": True,
            "message": "현재 상태 스크린샷 촬영 완료"
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

    result = check_current_state(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
