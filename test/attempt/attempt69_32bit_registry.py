"""
시도 69: 32비트 레지스트리 검색

64비트 Windows에서 32비트 애플리케이션의 COM 등록은
WOW6432Node 경로에 있음. 이 경로를 검색하여 FarPoint Spread 찾기
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, DWORD
import winreg
import os


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 69: 32비트 레지스트리 검색")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt69_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 포커스 설정
        left_spread.set_focus()
        time.sleep(0.5)

        # 참조 값 확인
        import pyperclip
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (복사): '{reference_value}'")

        print("\n=== 1. 32비트 CLSID 검색 (WOW6432Node) ===")

        found_clsids = []
        try:
            # WOW6432Node\CLSID 검색
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "WOW6432Node\\CLSID") as clsid_key:
                i = 0
                while i < 10000:
                    try:
                        clsid_name = winreg.EnumKey(clsid_key, i)
                        try:
                            # 설명 읽기
                            with winreg.OpenKey(clsid_key, clsid_name) as key:
                                try:
                                    value = winreg.QueryValue(key, "")
                                    if value and ("spread" in value.lower() or "farpoint" in value.lower() or "fpspread" in value.lower() or "fpuspread" in value.lower() or "spr32" in value.lower()):
                                        progid = None
                                        try:
                                            with winreg.OpenKey(key, "ProgID") as progid_key:
                                                progid = winreg.QueryValue(progid_key, "")
                                        except:
                                            pass

                                        # InprocServer32 경로
                                        ocx_path = None
                                        try:
                                            with winreg.OpenKey(key, "InprocServer32") as server_key:
                                                ocx_path = winreg.QueryValue(server_key, "")
                                        except:
                                            pass

                                        found_clsids.append({
                                            'clsid': clsid_name,
                                            'description': value,
                                            'progid': progid,
                                            'path': ocx_path
                                        })

                                        print(f"\n  {clsid_name}:")
                                        print(f"    설명: {value}")
                                        if progid:
                                            print(f"    ProgID: {progid}")
                                        if ocx_path:
                                            print(f"    경로: {ocx_path}")
                                except:
                                    pass
                        except:
                            pass
                        i += 1
                    except OSError:
                        break
        except FileNotFoundError:
            print("  ✗ WOW6432Node\\CLSID 키가 없음 (64비트 시스템이 아니거나 32비트 Python)")
        except Exception as e:
            print(f"  검색 오류: {e}")

        print(f"\n총 {len(found_clsids)}개 발견")

        print("\n=== 2. 32비트 TypeLib 검색 ===")

        found_typelibs = []
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "WOW6432Node\\TypeLib") as typelib_key:
                i = 0
                while i < 5000:
                    try:
                        typelib_guid = winreg.EnumKey(typelib_key, i)
                        try:
                            with winreg.OpenKey(typelib_key, typelib_guid) as guid_key:
                                j = 0
                                while j < 10:
                                    try:
                                        version = winreg.EnumKey(guid_key, j)
                                        try:
                                            desc = winreg.QueryValue(guid_key, version)
                                            if desc and ("spread" in desc.lower() or "farpoint" in desc.lower() or "fpspread" in desc.lower()):
                                                found_typelibs.append({
                                                    'guid': typelib_guid,
                                                    'version': version,
                                                    'description': desc
                                                })
                                                print(f"\n  {desc}")
                                                print(f"    GUID: {typelib_guid}")
                                                print(f"    버전: {version}")
                                        except:
                                            pass
                                        j += 1
                                    except OSError:
                                        break
                        except:
                            pass
                        i += 1
                    except OSError:
                        break
        except FileNotFoundError:
            print("  ✗ WOW6432Node\\TypeLib 키가 없음")
        except Exception as e:
            print(f"  검색 오류: {e}")

        print(f"\n총 {len(found_typelibs)}개 발견")

        # COM 객체 생성 시도
        if found_clsids:
            print("\n=== 3. COM 객체 생성 시도 ===")
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()

            for info in found_clsids[:3]:  # 처음 3개만 시도
                progid = info['progid']
                clsid = info['clsid']

                if progid:
                    print(f"\n시도: {progid}")
                    try:
                        obj = win32com.client.Dispatch(progid)
                        print(f"  ✓ COM 객체 생성 성공")

                        # 타입 정보 얻기
                        try:
                            obj_typed = win32com.client.gencache.EnsureDispatch(progid)
                            print("  사용 가능한 속성/메서드:")
                            count = 0
                            for attr in dir(obj_typed):
                                if not attr.startswith('_'):
                                    print(f"    - {attr}")
                                    count += 1
                                    if count > 20:  # 처음 20개만
                                        print("    ...")
                                        break
                        except Exception as e:
                            print(f"  타입 정보 조회 실패: {e}")

                    except Exception as e:
                        print(f"  ✗ 실패: {e}")

        # HWND로부터 직접 COM 객체 얻기 시도
        print("\n=== 4. HWND로부터 IDispatch 얻기 (WM_GETOBJECT) ===")

        import win32api
        import win32gui

        WM_GETOBJECT = 0x003D
        OBJID_NATIVEOM = 0xFFFFFFF0

        result = win32api.SendMessage(hwnd, WM_GETOBJECT, 0, OBJID_NATIVEOM)
        print(f"WM_GETOBJECT 결과: {result} (0x{result:08X})")

        if result:
            print("  ✓ 응답 있음 - ObjectFromLresult 시도")

            oleacc = windll.oleacc

            class GUID(Structure):
                _fields_ = [
                    ("Data1", DWORD),
                    ("Data2", c_ushort),
                    ("Data3", c_ushort),
                    ("Data4", c_ubyte * 8)
                ]

            IID_IDispatch = GUID(
                0x00020400, 0x0000, 0x0000,
                (c_ubyte * 8)(0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)
            )

            obj_ptr = c_void_p()
            hr = oleacc.ObjectFromLresult(
                result,
                byref(IID_IDispatch),
                0,
                byref(obj_ptr)
            )

            print(f"  ObjectFromLresult: 0x{hr:08X}")

            if hr == 0 and obj_ptr.value:
                print(f"  ✓ IDispatch 획득: 0x{obj_ptr.value:08X}")

                try:
                    import pythoncom
                    import win32com.client

                    disp = pythoncom.ObjectFromLong(obj_ptr.value)
                    wrapped = win32com.client.Dispatch(disp)

                    print(f"  ✓ COM 객체 래핑 성공")
                    print(f"    타입: {type(wrapped)}")

                    # 주요 메서드 시도
                    methods_to_try = [
                        ("ActiveSheet", None),
                        ("GetText", (1, 1)),  # row=1, col=1
                        ("Col", None),
                        ("Row", None),
                        ("Text", None),
                        ("Value", None),
                    ]

                    for method_name, args in methods_to_try:
                        try:
                            if args:
                                result = getattr(wrapped, method_name)(*args)
                            else:
                                result = getattr(wrapped, method_name, None)

                            if result is not None:
                                print(f"    ✓ {method_name}: {result}")
                        except Exception as e:
                            pass

                except Exception as e:
                    print(f"  ✗ 래핑 실패: {e}")
        else:
            print("  ✗ 응답 없음")

        capture_func("attempt69_01_complete.png")

        return {
            "success": False,
            "message": f"32비트 레지스트리 검색 완료: CLSID {len(found_clsids)}개, TypeLib {len(found_typelibs)}개"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
