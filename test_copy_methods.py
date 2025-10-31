"""
복사 메서드 테스트 실행 스크립트
"""
from pywinauto import application
import sys

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("="*70)
    print("fpUSpread80 복사 메서드 테스트")
    print("="*70)

    # 연결
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        hwnd = dlg.handle

        print(f"\n✓ 사원등록 윈도우 연결 성공 (HWND: 0x{hwnd:08X})\n")

    except Exception as e:
        print(f"\n✗ 연결 실패: {e}")
        print("→ 사원등록 프로그램이 실행 중인지 확인하세요")
        return

    # capture 함수 생성
    from test.capture import capture_window
    def capture_func(filename):
        return capture_window(hwnd, filename)

    # 테스트 실행
    from test.attempt.attempt54_test_copy_methods import run as attempt54
    result = attempt54(dlg, capture_func)

    print(f"\n{'='*70}")
    print(f"최종 결과: {result}")
    print(f"{'='*70}")

if __name__ == "__main__":
    main()
