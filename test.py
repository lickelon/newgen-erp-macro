"""
사원등록 자동화 테스트 메인 스크립트
"""
from pywinauto import application
import sys
from test.capture import capture_window

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("="*70)
    print("사원등록 자동화 테스트")
    print("="*70)

    # 연결
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        hwnd = dlg.handle

        print(f"\n✓ 사원등록 윈도우 연결 성공 (HWND: {hwnd})\n")

    except Exception as e:
        print(f"\n✗ 연결 실패: {e}")
        return

    # capture 함수 생성
    def capture_func(filename):
        capture_window(hwnd, filename)

    # 시도 85: COM Dispatch 방식
    print("\n⚠️  시도 85 실행 중...")
    from test.attempt.attempt85_com_dispatch import run as attempt85

    result = attempt85(dlg, capture_func)

    print("\n" + "="*70)
    print(f"결과: {result['message']}")
    print("="*70)

    if result["success"]:
        print("\n✅ 테스트 완료!")
    else:
        print("\n⚠️  테스트 실패.")

if __name__ == "__main__":
    main()
