"""
연말정산 자동화 테스트 메인 스크립트
"""
from pywinauto import application
import sys
from test.capture import capture_window

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("="*70)
    print("연말정산 자동화 테스트")
    print("="*70)

    # 연결
    try:
        app = application.Application(backend="win32")
        app.connect(title="연말정산추가자료입력")
        dlg = app.window(title="연말정산추가자료입력")
        hwnd = dlg.handle

        print(f"\n✓ 연말정산 윈도우 연결 성공 (HWND: {hwnd})\n")

    except Exception as e:
        print(f"\n✗ 연결 실패: {e}")
        return

    # capture 함수 생성
    def capture_func(filename):
        capture_window(hwnd, filename)

    # 시도 1: 탭 컨트롤 자식 요소 클릭
    from test.attempt.attempt01_click_children import run as attempt01

    result = attempt01(dlg, capture_func)

    print("\n" + "="*70)
    print(f"결과: {result['message']}")
    print("="*70)

    if result["success"]:
        print("\n✅ 테스트 완료! test/image/ 폴더의 스크린샷을 확인하세요.")
    else:
        print("\n⚠️  테스트 실패. 다음 시도를 준비하세요.")

if __name__ == "__main__":
    main()
