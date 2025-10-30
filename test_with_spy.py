"""
Spy++ 실시간 모니터링과 함께 탭 자동화 테스트

사용 방법:
1. 사원등록 프로그램 실행
2. Spy++ 실행 및 메시지 로그 시작 (Spy → Messages)
3. 이 스크립트 실행: uv run python test_with_spy.py
"""
import sys
import time
from tab_automation import TabAutomation

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')


def print_separator(char="=", length=70):
    """구분선 출력"""
    print(char * length)


def main():
    print_separator()
    print("Spy++ 실시간 모니터링 테스트")
    print_separator()

    print("\n📋 준비사항:")
    print("1. 사원등록 프로그램이 실행되어 있나요?")
    print("2. Spy++를 실행했나요?")
    print("3. Spy++에서 메시지 로그를 시작했나요? (Spy → Messages → Ctrl+M)")
    print("\n")
    print("Spy++ 설정 팁:")
    print("- 타겟: 사원등록 윈도우 또는 탭 컨트롤")
    print("- 필터: WM_LBUTTONDOWN, WM_LBUTTONUP, WM_NOTIFY 선택")
    print("- WM_PAINT, WM_TIMER는 제외 (너무 많음)")

    input("\n준비 완료! Enter를 누르면 자동화 스크립트를 실행합니다...")

    print_separator("-")
    print("🤖 자동화 시작")
    print_separator("-")

    try:
        # 연결
        print("\n[1/4] 사원등록 윈도우 연결 중...")
        tab_auto = TabAutomation()
        tab_auto.connect()
        print("✓ 연결 성공")

        time.sleep(1)

        # 탭 1: 부양가족정보
        print("\n[2/4] '부양가족정보' 탭 선택 중...")
        print("👀 Spy++를 확인하세요!")
        print("   - WM_LBUTTONDOWN (lparam: 0x000F0096 = x:150, y:15)")
        print("   - WM_LBUTTONUP (lparam: 0x000F0096)")
        print("   - WM_NOTIFY (탭 변경 알림)")
        tab_auto.select_tab("부양가족정보")
        print("✓ 탭 선택 완료")

        time.sleep(2)

        # 탭 2: 소득자료
        print("\n[3/4] '소득자료' 탭 선택 중...")
        print("👀 Spy++를 확인하세요!")
        print("   - WM_LBUTTONDOWN (lparam: 0x000F00FA = x:250, y:15)")
        print("   - WM_LBUTTONUP (lparam: 0x000F00FA)")
        print("   - WM_NOTIFY (탭 변경 알림)")
        tab_auto.select_tab("소득자료")
        print("✓ 탭 선택 완료")

        time.sleep(2)

        # 탭 3: 기본사항 (원위치)
        print("\n[4/4] '기본사항' 탭 선택 중 (원위치)...")
        print("👀 Spy++를 확인하세요!")
        print("   - WM_LBUTTONDOWN (lparam: 0x000F0032 = x:50, y:15)")
        print("   - WM_LBUTTONUP (lparam: 0x000F0032)")
        print("   - WM_NOTIFY (탭 변경 알림)")
        tab_auto.select_tab("기본사항")
        print("✓ 탭 선택 완료")

        print_separator("-")
        print("✅ 모든 테스트 완료!")
        print_separator("-")

    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return

    # 결과 분석
    print("\n📊 Spy++ 로그 분석")
    print_separator("-")
    print("Spy++ 로그에서 다음을 확인하세요:")
    print()
    print("✅ 체크리스트:")
    print("  [ ] WM_LBUTTONDOWN 메시지가 3번 보였나요?")
    print("  [ ] WM_LBUTTONUP 메시지가 3번 보였나요?")
    print("  [ ] lparam 좌표가 올바른가요?")
    print("      - 부양가족정보: x=150(0x96), y=15(0x0F)")
    print("      - 소득자료: x=250(0xFA), y=15(0x0F)")
    print("      - 기본사항: x=50(0x32), y=15(0x0F)")
    print("  [ ] 각 클릭 후 WM_NOTIFY가 발생했나요?")
    print("  [ ] 마우스 커서가 움직이지 않았나요?")
    print()
    print("💾 로그 저장:")
    print("  Spy++ → File → Save → log_with_automation.txt")
    print()
    print("📖 자세한 내용:")
    print("  docs/spy-realtime-monitoring.md 참조")
    print_separator()


if __name__ == "__main__":
    main()
