"""
시도 97: print_control_identifiers로 전체 구조 출력
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
    print("시도 97: print_control_identifiers 출력")
    print("="*60)

    try:
        # 급여자료입력 연결
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        print(f"\n급여자료입력 HWND: 0x{main_window.handle:08X}")

        # 전체 컨트롤 구조 출력
        print("\n" + "="*60)
        print("전체 컨트롤 식별자 출력 (파일 저장)")
        print("="*60)

        # 파일로 저장
        import sys
        from io import StringIO

        old_stdout = sys.stdout
        sys.stdout = StringIO()

        try:
            main_window.print_control_identifiers()
            output = sys.stdout.getvalue()
        finally:
            sys.stdout = old_stdout

        # 파일에 저장
        with open("test/control_identifiers.txt", "w", encoding="utf-8") as f:
            f.write(output)

        print(f"✓ 저장: test/control_identifiers.txt ({len(output)} bytes)")

        # 분납 관련 라인만 출력
        print("\n'분납' 포함 라인:")
        for line in output.split('\n'):
            if '분납' in line:
                print(f"  {line}")

        return {"success": True, "message": "컨트롤 식별자 출력 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
