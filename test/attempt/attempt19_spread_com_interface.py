"""
Attempt 19: fpUSpread80 COM 인터페이스 탐색

Farpoint Spread는 COM 컨트롤이므로 win32com으로 접근 가능
정확한 행/열 인덱스로 셀을 지정할 수 있음
"""
import win32com.client
import pythoncom
from pywinauto import application
import sys


def explore_spread_com(dlg, capture_func):
    """fpUSpread80 COM 인터페이스 탐색"""
    print("\n" + "=" * 60)
    print("Attempt 19: fpUSpread80 COM 인터페이스 탐색")
    print("=" * 60)

    try:
        capture_func("attempt19_00_initial.png")

        # fpUSpread80 찾기
        print("\n[1/4] fpUSpread80 컨트롤 찾기...")
        spread_controls = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            return {"success": False, "message": f"Spread 부족: {len(spread_controls)}"}

        # Spread #2 (왼쪽 목록)
        left_list = spread_controls[2]
        hwnd = left_list.handle

        print(f"  Spread #2 HWND: 0x{hwnd:08X}")

        # COM 인터페이스 접근 시도
        print("\n[2/4] COM 인터페이스 접근 시도...")

        # 방법 1: HWND로 IDispatch 얻기
        try:
            import ctypes
            from ctypes import POINTER, c_void_p, byref
            from comtypes import IUnknown

            print("  시도 1: ObjectFromLresult...")
            # ObjectFromLresult를 사용하여 IDispatch 얻기
            # 이건 복잡하므로 일단 스킵

        except Exception as e:
            print(f"  ✗ 방법 1 실패: {e}")

        # 방법 2: pywinauto wrapper에서 COM 객체 추출
        try:
            print("\n  시도 2: pywinauto wrapper...")
            # left_list.wrapper_object()
            if hasattr(left_list, 'iface_IDispatch'):
                dispatch = left_list.iface_IDispatch
                print(f"  ✓ IDispatch 발견: {dispatch}")
            else:
                print("  ✗ IDispatch 없음")
        except Exception as e:
            print(f"  ✗ 방법 2 실패: {e}")

        # 방법 3: 컨트롤 속성 탐색
        print("\n[3/4] 컨트롤 속성 탐색...")
        try:
            attrs = dir(left_list)
            interesting_attrs = [
                attr for attr in attrs
                if not attr.startswith('_') and
                ('com' in attr.lower() or 'dispatch' in attr.lower() or
                 'ole' in attr.lower() or 'iface' in attr.lower())
            ]
            print(f"  관심 속성: {interesting_attrs}")

            # 각 속성 조사
            for attr in interesting_attrs[:5]:  # 처음 5개만
                try:
                    value = getattr(left_list, attr)
                    print(f"    {attr}: {type(value)}")
                except Exception as e:
                    print(f"    {attr}: 접근 실패 - {e}")

        except Exception as e:
            print(f"  ✗ 속성 탐색 실패: {e}")

        # 방법 4: 메시지로 정보 얻기
        print("\n[4/4] SendMessage로 정보 얻기...")
        import win32api
        import win32con

        # 일부 ActiveX 컨트롤은 특수 메시지 지원
        try:
            # 행/열 개수 얻기 시도 (추측)
            result = win32api.SendMessage(hwnd, win32con.WM_USER + 100, 0, 0)
            print(f"  WM_USER+100: {result}")

            result = win32api.SendMessage(hwnd, win32con.WM_USER + 101, 0, 0)
            print(f"  WM_USER+101: {result}")

            # fpSpread는 자체 메시지를 가지고 있을 수 있음
            # 문서 없이는 추측하기 어려움

        except Exception as e:
            print(f"  ✗ 메시지 실패: {e}")

        capture_func("attempt19_01_explored.png")

        print("\n" + "=" * 60)
        print("COM 인터페이스 직접 접근은 어려움")
        print("대안: Tab 키 네비게이션 또는 정확한 좌표 측정")
        print("=" * 60)

        return {
            "success": False,
            "message": "COM 인터페이스 접근 불가, 대안 필요"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }


if __name__ == "__main__":
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    # 경로 설정
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    # 연결
    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    # 실행
    result = explore_spread_com(dlg, capture_func)
    print(f"\n최종 결과: {result}")
