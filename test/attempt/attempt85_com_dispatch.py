"""
시도 85: pywin32 COM Dispatch를 통한 FarPoint Spread 접근

win32com.client.Dispatch로 실행 중인 FarPoint Spread COM 객체에 접근 시도
"""
import time
import win32gui
import win32con
import win32api


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 85: pywin32 COM Dispatch")
    print("="*60)

    try:
        import win32com.client
        print("✓ win32com 라이브러리 로드 성공")
    except ImportError:
        return {
            "success": False,
            "message": "pywin32가 설치되지 않았습니다! pip install pywin32"
        }

    try:
        # 초기 상태 캡처
        capture_func("attempt85_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        spread_hwnd = left_spread.handle

        print(f"\n왼쪽 스프레드 HWND: 0x{spread_hwnd:08X}")

        print("\n=== ProgID로 새 인스턴스 생성 시도 ===")

        # FarPoint Spread의 가능한 ProgID들
        progids = [
            "FarPoint.Spread.8",
            "FarPoint.Spread",
            "FPSpread.vaSpread",
            "fpSpread.fpSpread",
        ]

        for progid in progids:
            try:
                print(f"\n  시도: {progid}")
                spread = win32com.client.Dispatch(progid)
                print(f"    ✓ Dispatch 성공")
                print(f"    (새 인스턴스이므로 기존 데이터 없음)")
            except Exception as e:
                print(f"    ✗ 실패: {e}")

        capture_func("attempt85_02_complete.png")

        return {
            "success": False,
            "message": """
COM 접근 실패

시도한 방법:
- ProgID로 새 인스턴스 생성

모두 실패했습니다. fpUSpread80은 표준 COM 인터페이스를
노출하지 않는 것으로 보입니다.
"""
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
