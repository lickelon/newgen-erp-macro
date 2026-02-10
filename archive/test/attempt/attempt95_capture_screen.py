"""
시도 95: 화면 캡처
"""
from pywinauto import Application

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 스크린샷 함수

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 95: 화면 캡처")
    print("="*60)

    try:
        # 급여자료입력 창에 연결
        print("\n급여자료입력 창 캡처 중...")
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        # 캡처
        capture_func("attempt95_salary_window.png")
        print("✓ 저장: test/image/attempt95_salary_window.png")

        return {"success": True, "message": "화면 캡처 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
