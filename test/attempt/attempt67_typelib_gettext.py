"""
시도 67: TypeLib 로드 후 GetText 메서드 호출

레지스트리에서 fpUSpread80의 CLSID/ProgID를 찾아서
COM 객체를 생성하고 ActiveSheet.GetText() 메서드 호출
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, DWORD
import winreg
import pythoncom
import win32com.client


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 67: TypeLib 로드 후 GetText 메서드 호출")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt67_00_initial.png")

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

        print("\n=== 레지스트리에서 FarPoint/Spread CLSID 찾기 ===")

        found_clsids = []

        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID") as clsid_key:
                i = 0
                while i < 5000:  # 처음 5000개만
                    try:
                        clsid_name = winreg.EnumKey(clsid_key, i)

                        try:
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

                                        found_clsids.append({
                                            'clsid': clsid_name,
                                            'description': value,
                                            'progid': progid
                                        })

                                        print(f"  {clsid_name}: {value}")
                                        if progid:
                                            print(f"    ProgID: {progid}")

                                except:
                                    pass
                        except:
                            pass

                        i += 1
                    except OSError:
                        break

        except Exception as e:
            print(f"레지스트리 검색 실패: {e}")

        if not found_clsids:
            print("✗ FarPoint/Spread 관련 CLSID를 찾을 수 없음")

        print(f"\n총 {len(found_clsids)}개 발견")

        print("\n=== COM 객체 생성 시도 ===")

        pythoncom.CoInitialize()

        # 찾은 ProgID/CLSID로 COM 객체 생성 시도
        for info in found_clsids:
            progid = info['progid']
            clsid = info['clsid']

            if progid:
                print(f"\n시도: {progid}")
                try:
                    # ProgID로 COM 객체 생성
                    obj = win32com.client.Dispatch(progid)
                    print(f"✓ COM 객체 생성 성공")

                    # 사용 가능한 메서드/속성 확인
                    try:
                        # EnsureDispatch로 타입 정보 얻기
                        obj_typed = win32com.client.gencache.EnsureDispatch(progid)
                        print("  사용 가능한 속성/메서드:")
                        for attr in dir(obj_typed):
                            if not attr.startswith('_'):
                                print(f"    - {attr}")
                    except:
                        pass

                except Exception as e:
                    print(f"✗ 실패: {e}")

        print("\n=== HWND로부터 COM 객체 얻기 시도 ===")

        # WM_GETOBJECT 사용
        import win32api
        WM_GETOBJECT = 0x003D
        OBJID_NATIVEOM = 0xFFFFFFF0

        result = win32api.SendMessage(hwnd, WM_GETOBJECT, 0, OBJID_NATIVEOM)
        print(f"WM_GETOBJECT 응답: {result}")

        if result:
            # ObjectFromLresult 시도
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

            print(f"ObjectFromLresult: 0x{hr:08X}")

            if hr == 0 and obj_ptr.value:
                print(f"✓ IDispatch 획득: 0x{obj_ptr.value:08X}")

                # PyIDispatch로 래핑
                try:
                    disp = pythoncom.ObjectFromLong(obj_ptr.value)
                    wrapped = win32com.client.Dispatch(disp)

                    print("✓ COM 객체 래핑 성공")
                    print(f"  타입: {type(wrapped)}")

                    # 속성/메서드 시도
                    methods_to_try = [
                        "ActiveSheet",
                        "ActiveSheetView",
                        "Sheets",
                        "GetText",
                        "Col",
                        "Row",
                    ]

                    for method in methods_to_try:
                        try:
                            result = getattr(wrapped, method, None)
                            if result is not None:
                                print(f"  ✓ {method}: {result}")
                        except Exception as e:
                            pass

                except Exception as e:
                    print(f"✗ 래핑 실패: {e}")

        capture_func("attempt67_01_complete.png")

        return {
            "success": False,
            "message": "TypeLib/COM 객체 접근 실패"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
