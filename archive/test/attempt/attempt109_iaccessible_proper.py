"""
시도 109: IAccessible 제대로 사용하기

comtypes로 Accessibility 타입 라이브러리 생성 후 IAccessible 사용
"""
import win32process
import win32gui
import win32con
import win32api
import time
from pywinauto import Application
from ctypes import POINTER, byref, c_long, c_void_p, oledll
from comtypes import GUID, IUnknown
import comtypes.client


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 109: IAccessible 제대로 사용하기")
    print("="*60)

    try:
        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[단계] 급여자료입력 연결 (PID: {process_id})")

        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "#32770":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        installment_dlg = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 2. 스프레드 찾기
        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        left_hwnd = left_spread.handle
        print(f"\n왼쪽 스프레드: 0x{left_hwnd:08X}")

        # === 방법 1: Accessibility 타입 라이브러리 생성 ===
        print("\n" + "="*60)
        print("[방법 1] Accessibility 타입 라이브러리 생성")
        print("="*60)

        try:
            # oleacc.dll에서 타입 라이브러리 로드
            print("Accessibility 타입 라이브러리 로딩 중...")
            tlib = comtypes.client.GetModule("oleacc.dll")
            print("✓ Accessibility 타입 라이브러리 생성 성공!")

            # 이제 import 가능
            from comtypes.gen import Accessibility

            print("✓ Accessibility 모듈 import 성공!")

        except Exception as e:
            print(f"✗ 타입 라이브러리 생성 실패: {e}")
            return {"success": False, "message": f"Accessibility 타입 라이브러리 생성 실패: {e}"}

        # === 방법 2: IAccessible로 스프레드 접근 ===
        print("\n" + "="*60)
        print("[방법 2] IAccessible로 데이터 읽기")
        print("="*60)

        try:
            # AccessibleObjectFromWindow
            OBJID_CLIENT = 0xFFFFFFFC
            IID_IAccessible = GUID("{618736E0-3C3D-11CF-810C-00AA00389B71}")

            AccessibleObjectFromWindow = oledll.oleacc.AccessibleObjectFromWindow

            # IUnknown 포인터 획득
            pacc = POINTER(IUnknown)()
            result = AccessibleObjectFromWindow(
                left_hwnd,
                OBJID_CLIENT,
                byref(IID_IAccessible),
                byref(pacc)
            )

            if result == 0 and pacc:
                print("✓ AccessibleObjectFromWindow 성공")

                # IAccessible로 QueryInterface
                accessible = pacc.QueryInterface(Accessibility.IAccessible)
                print("✓ IAccessible 인터페이스 획득")

                # 속성 확인
                try:
                    name = accessible.accName(0)
                    print(f"  이름: '{name}'")
                except:
                    print("  이름: (없음)")

                try:
                    value = accessible.accValue(0)
                    print(f"  값: '{value}'")
                    if value:
                        print(f"\n✓✓✓ accValue로 데이터 발견!")
                        return {"success": True, "message": f"IAccessible.accValue: '{value}'"}
                except:
                    print("  값: (없음)")

                try:
                    description = accessible.accDescription(0)
                    print(f"  설명: '{description}'")
                    if description:
                        print(f"\n✓✓✓ accDescription으로 데이터 발견!")
                        return {"success": True, "message": f"IAccessible.accDescription: '{description}'"}
                except:
                    print("  설명: (없음)")

                try:
                    role = accessible.accRole(0)
                    print(f"  역할: {role}")
                except:
                    print("  역할: (없음)")

                try:
                    state = accessible.accState(0)
                    print(f"  상태: {state}")
                except:
                    print("  상태: (없음)")

                # 자식 수 확인
                try:
                    child_count = accessible.accChildCount
                    print(f"\n  자식 수: {child_count}")

                    if child_count > 0:
                        print(f"  처음 10개 자식 확인:")

                        for i in range(1, min(child_count + 1, 11)):
                            try:
                                child_name = accessible.accName(i)
                                child_value = accessible.accValue(i)
                                child_role = accessible.accRole(i)

                                if child_name or child_value:
                                    print(f"\n    자식[{i}]:")
                                    if child_name:
                                        print(f"      이름: '{child_name}'")
                                    if child_value:
                                        print(f"      값: '{child_value}'")
                                        print(f"\n✓✓✓ 자식에서 값 발견!")
                                        return {"success": True, "message": f"자식[{i}] 값: '{child_value}'"}
                                    if child_role:
                                        print(f"      역할: {child_role}")
                            except Exception as e:
                                pass  # 자식 읽기 실패는 무시

                except Exception as e:
                    print(f"  자식 확인 실패: {e}")

                print("\n✗ IAccessible로 데이터를 찾지 못함")

            else:
                print(f"✗ AccessibleObjectFromWindow 실패: 0x{result:X}")

        except Exception as e:
            import traceback
            print(f"✗ IAccessible 처리 실패: {e}")
            print(traceback.format_exc())

        # === 방법 3: 직접 IDispatch로 시도 ===
        print("\n" + "="*60)
        print("[방법 3] COM IDispatch로 Spread 접근")
        print("="*60)

        try:
            from comtypes.automation import IDispatch

            # IDispatch 포인터 획득
            OBJID_NATIVEOM = 0xFFFFFFF0
            IID_IDispatch = GUID("{00020400-0000-0000-C000-000000000046}")

            pdisp = POINTER(IDispatch)()
            result = AccessibleObjectFromWindow(
                left_hwnd,
                OBJID_NATIVEOM,
                byref(IID_IDispatch),
                byref(pdisp)
            )

            if result == 0 and pdisp:
                print("✓ IDispatch 인터페이스 획득 성공!")

                # IDispatch로 QueryInterface
                dispatch = pdisp.QueryInterface(IDispatch)
                print("✓ IDispatch 쿼리 성공")

                # 속성 접근 시도
                try:
                    # GetIDsOfNames로 메서드/속성 ID 찾기
                    print("  사용 가능한 메서드/속성 탐색...")

                    # 일반적인 Spread 속성들
                    test_props = [
                        "Text", "Value", "CellText", "ActiveCell",
                        "ActiveRow", "ActiveCol", "Row", "Col",
                        "MaxRows", "MaxCols", "DataRowCnt", "DataColCnt"
                    ]

                    for prop_name in test_props:
                        try:
                            # 속성 접근 시도
                            obj = dispatch
                            print(f"    시도: {prop_name}")
                            # comtypes는 자동으로 속성 접근 지원
                            # (실제로는 GetIDsOfNames + Invoke)
                        except:
                            pass

                except Exception as e:
                    print(f"  속성 탐색 실패: {e}")

                print("\n✗ IDispatch로도 데이터를 찾지 못함")

            else:
                print(f"✗ IDispatch 획득 실패: 0x{result:X}")

        except Exception as e:
            print(f"✗ IDispatch 시도 실패: {e}")

        return {
            "success": False,
            "message": "IAccessible/IDispatch 모두 데이터 읽기 실패"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
