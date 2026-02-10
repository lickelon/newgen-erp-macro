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

    # 시도 117: 순서 기반 자동 입력
    print("\n⚠️  시도 117 실행 중...")
    from test.attempt.attempt117_sequential_input import run as attempt117

    result = attempt117(None, None)

    print("\n" + "="*70)
    print(f"결과: {result['message']}")
    print("="*70)

    if result["success"]:
        print("\n✅ 테스트 완료!")
    else:
        print("\n⚠️  테스트 실패.")

if __name__ == "__main__":
    main()
