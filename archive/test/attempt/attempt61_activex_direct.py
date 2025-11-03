"""
시도 61: Farpoint Spread ActiveX 직접 접근

HWND로부터 ActiveX 컨트롤의 IDispatch 인터페이스를 얻어서
COM 메서드를 직접 호출하기
"""
import time
import win32com.client
import pythoncom
from ctypes import *
from comtypes import GUID, IUnknown, COMMETHOD, HRESULT, POINTER
import comtypes.client


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 61: Farpoint Spread ActiveX 직접 접근")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt61_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 포커스 설정하고 참조 값 확인
        left_spread.set_focus()
        time.sleep(0.5)

        import pyperclip
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (복사): '{reference_value}'")

        # 현재 활성 셀 위치 확인 (HOME으로 가서 위치 파악)
        left_spread.type_keys("{HOME}", pause=0.05)
        time.sleep(0.2)
        left_spread.type_keys("{RIGHT}", pause=0.05)  # 체크박스 건너뛰기
        time.sleep(0.2)
        pyperclip.copy("BEFORE2")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        col1_value = pyperclip.paste()
        print(f"첫 번째 컬럼 값: '{col1_value}'")

        print("\n=== 방법 1: comtypes로 IDispatch 얻기 ===")
        try:
            pythoncom.CoInitialize()

            # IDispatch GUID
            IID_IDispatch = GUID("{00020400-0000-0000-C000-000000000046}")

            oleacc = windll.oleacc
            obj_ptr = POINTER(IUnknown)()

            result = oleacc.AccessibleObjectFromWindow(
                hwnd,
                0,  # OBJID_CLIENT
                byref(IID_IDispatch),
                byref(obj_ptr)
            )

            if result == 0 and obj_ptr:
                print(f"✓ IDispatch 포인터 획득: {obj_ptr}")

                # comtypes로 래핑
                try:
                    disp = obj_ptr.QueryInterface(comtypes.automation.IDispatch)
                    print(f"✓ IDispatch 인터페이스: {disp}")

                    # 타입 정보 확인
                    try:
                        typeinfo = disp.GetTypeInfo(0)
                        typeattr = typeinfo.GetTypeAttr()
                        print(f"  타입 이름: {typeinfo.GetDocumentation(-1)[0]}")
                        print(f"  함수 개수: {typeattr.cFuncs}")
                        print(f"  변수 개수: {typeattr.cVars}")

                        # 모든 메서드/속성 출력
                        print("\n사용 가능한 메서드/속성:")
                        for i in range(typeattr.cFuncs):
                            funcdesc = typeinfo.GetFuncDesc(i)
                            name = typeinfo.GetNames(funcdesc.memid)[0]
                            print(f"  - {name}")

                    except Exception as e:
                        print(f"  타입 정보 읽기 실패: {e}")

                    # 일반적인 Spread 메서드 시도
                    methods_to_try = [
                        ("GetText", [1, 1]),  # col, row (1-based)
                        ("GetText", [2, 1]),
                        ("get_Text", [1, 1]),
                        ("Text", [1, 1]),
                        ("GetValue", [1, 1]),
                        ("get_Value", [1, 1]),
                        ("Value", [1, 1]),
                        ("GetData", [1, 1]),
                        ("Cell", [1, 1]),
                        ("ActiveCell", []),
                        ("Col", []),
                        ("Row", []),
                    ]

                    print("\n메서드 호출 시도:")
                    for method_name, args in methods_to_try:
                        try:
                            # GetIDsOfNames로 메서드 ID 얻기
                            method_id = disp.GetIDsOfNames([method_name])[0]
                            print(f"  {method_name}: ID={method_id}")

                            # Invoke 호출
                            result = disp.Invoke(method_id, 0, 1, *args)  # DISPATCH_METHOD
                            print(f"    결과: {result}")

                        except Exception as e:
                            # 에러 메시지가 "이름을 찾을 수 없음"이 아니면 출력
                            err_msg = str(e)
                            if "Unknown name" not in err_msg and "이름" not in err_msg:
                                print(f"  {method_name}: {err_msg[:50]}")

                except Exception as e:
                    import traceback
                    print(f"✗ IDispatch 래핑 실패: {e}")
                    traceback.print_exc()

            else:
                print(f"✗ AccessibleObjectFromWindow 실패: {result}")

        except Exception as e:
            import traceback
            print(f"✗ 방법 1 실패: {e}")
            traceback.print_exc()

        print("\n=== 방법 2: win32com.client.Dispatch ===")
        try:
            pythoncom.CoInitialize()

            # HWND로부터 COM 객체 얻기 (다른 방법)
            import ctypes
            from ctypes.wintypes import HWND

            # ObjectFromLresult 시도
            oleacc = ctypes.windll.oleacc

            # IDispatch GUID
            IID_IDispatch = GUID("{00020400-0000-0000-C000-000000000046}")

            # 메시지를 보내서 IDispatch 포인터 얻기
            WM_GETOBJECT = 0x003D
            OBJID_CLIENT = 0xFFFFFFFC

            import win32api
            lresult = win32api.SendMessage(hwnd, WM_GETOBJECT, 0, OBJID_CLIENT)

            if lresult:
                print(f"WM_GETOBJECT 응답: {lresult}")

                # ObjectFromLresult 호출
                obj_ptr = POINTER(IUnknown)()
                result = oleacc.ObjectFromLresult(
                    lresult,
                    byref(IID_IDispatch),
                    0,
                    byref(obj_ptr)
                )

                if result == 0 and obj_ptr:
                    print(f"✓ ObjectFromLresult 성공")
                    # 위와 동일한 처리...
                else:
                    print(f"✗ ObjectFromLresult 실패: {result}")
            else:
                print("✗ WM_GETOBJECT 응답 없음")

        except Exception as e:
            import traceback
            print(f"✗ 방법 2 실패: {e}")
            traceback.print_exc()

        print("\n=== 방법 3: ProgID로 타입 라이브러리 탐색 ===")
        try:
            import winreg

            print("등록된 Farpoint/Spread ActiveX 검색:")

            # CLSID 레지스트리 검색
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID") as clsid_key:
                i = 0
                while True:
                    try:
                        clsid_name = winreg.EnumKey(clsid_key, i)

                        try:
                            with winreg.OpenKey(clsid_key, clsid_name) as key:
                                try:
                                    value = winreg.QueryValue(key, "")
                                    if "spread" in value.lower() or "farpoint" in value.lower():
                                        print(f"  {clsid_name}: {value}")

                                        # ProgID 확인
                                        try:
                                            with winreg.OpenKey(key, "ProgID") as progid_key:
                                                progid = winreg.QueryValue(progid_key, "")
                                                print(f"    ProgID: {progid}")
                                        except:
                                            pass

                                except:
                                    pass
                        except:
                            pass

                        i += 1
                        if i > 1000:  # 안전장치
                            break
                    except OSError:
                        break

        except Exception as e:
            print(f"✗ 방법 3 실패: {e}")

        capture_func("attempt61_01_complete.png")

        return {
            "success": False,
            "message": "ActiveX 직접 접근 실패 - ProgID 또는 타입 라이브러리 필요"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
