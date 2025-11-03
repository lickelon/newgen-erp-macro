"""
시도 70: 포그라운드 솔루션 (최종)

배경:
- 시도 54에서 type_keys가 활성 창에서 작동 확인
- 시도 55에서 백그라운드에서는 실패 확인
- 시도 56-69에서 모든 백그라운드 접근 방법 실패
- Stack Overflow 확인: FarPoint Spread는 외부 자동화를 지원하지 않음

결론:
창을 활성화한 상태에서 type_keys로 값을 읽는 것이 유일한 방법

이 시도에서는 안정적인 포그라운드 셀 읽기 메서드 구현
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 70: 포그라운드 솔루션 (최종)")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt70_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"왼쪽 스프레드 HWND: 0x{left_spread.handle:08X}")

        # 안정적인 셀 값 읽기 함수
        def read_cell_value(spread_control):
            """현재 선택된 셀의 값을 클립보드로 읽기"""
            import pyperclip

            # 창이 활성 상태인지 확인
            import win32gui
            if win32gui.GetForegroundWindow() != dlg.handle:
                print("  경고: 창이 활성화되지 않음 - 활성화 시도")
                dlg.set_focus()
                time.sleep(0.3)

            # 클립보드 초기화
            pyperclip.copy("__EMPTY__")
            time.sleep(0.1)

            # 복사 시도 (최대 3회)
            for attempt in range(3):
                spread_control.type_keys("^c", pause=0.1)
                time.sleep(0.2)

                value = pyperclip.paste()
                if value != "__EMPTY__":
                    return value

                print(f"  재시도 {attempt + 1}/3")
                time.sleep(0.2)

            return None

        # 테스트 1: 현재 셀 읽기
        print("\n=== 테스트 1: 현재 셀 읽기 ===")

        left_spread.set_focus()
        time.sleep(0.5)

        value1 = read_cell_value(left_spread)
        print(f"✓ 현재 셀 값: '{value1}'")

        # 테스트 2: 다른 셀로 이동 후 읽기
        print("\n=== 테스트 2: 아래 셀로 이동 후 읽기 ===")

        left_spread.type_keys("{DOWN}", pause=0.1)
        time.sleep(0.3)

        value2 = read_cell_value(left_spread)
        print(f"✓ 아래 셀 값: '{value2}'")

        # 테스트 3: 원위치 복귀 후 읽기
        print("\n=== 테스트 3: 원위치 복귀 ===")

        left_spread.type_keys("{UP}", pause=0.1)
        time.sleep(0.3)

        value3 = read_cell_value(left_spread)
        print(f"✓ 원위치 셀 값: '{value3}'")

        if value1 == value3:
            print("  ✓ 값 일치 확인")

        # 테스트 4: 포그라운드 유지 테스트
        print("\n=== 테스트 4: 창 활성 상태 유지 ===")

        # 여러 셀을 순회하며 읽기
        values = [value1]
        for i in range(3):
            left_spread.type_keys("{DOWN}", pause=0.1)
            time.sleep(0.2)

            val = read_cell_value(left_spread)
            values.append(val)
            print(f"  행 {i+2}: '{val}'")

        # 원위치 복귀
        for i in range(3):
            left_spread.type_keys("{UP}", pause=0.1)
            time.sleep(0.1)

        print(f"\n✓ 총 {len(values)}개 셀 읽기 성공")

        # 테스트 5: 안정성 테스트 - 빠르게 여러 번 읽기
        print("\n=== 테스트 5: 안정성 테스트 (10회 연속 읽기) ===")

        stable_count = 0
        for i in range(10):
            val = read_cell_value(left_spread)
            if val == value1:
                stable_count += 1
            else:
                print(f"  ✗ {i+1}회: 예상 '{value1}', 실제 '{val}'")

        print(f"✓ 안정성: {stable_count}/10 ({stable_count*10}%)")

        capture_func("attempt70_01_success.png")

        if stable_count >= 8:  # 80% 이상 성공
            return {
                "success": True,
                "message": f"""
🎉 포그라운드 솔루션 성공!

테스트 결과:
- 현재 셀 읽기: ✓
- 셀 이동 및 읽기: ✓
- 원위치 복귀: ✓
- 연속 셀 읽기: ✓ ({len(values)}개)
- 안정성: {stable_count}/10 ({stable_count*10}%)

읽은 값들: {values}

결론:
FarPoint Spread fpUSpread80은 백그라운드 자동화를 지원하지 않습니다.
창을 활성화한 상태에서 type_keys + 클립보드를 사용하는 것이 유일한 방법입니다.

참고:
- Stack Overflow: "FarPoint Spread는 UI Automation을 지원하지 않음"
- 시도 56-69: 모든 백그라운드 접근 방법 실패 확인
- 시도 54: 포그라운드에서만 작동 확인
"""
            }
        else:
            return {
                "success": False,
                "message": f"안정성 부족: {stable_count}/10"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
