"""
시도 65: BSTR 추출 수정

VARIANT에서 BSTR을 올바르게 추출하여 값 읽기
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, DWORD


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 65: BSTR 추출 수정")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt65_00_initial.png")

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

        print("\n=== IAccessible 획득 ===")

        # GUID 정의
        class GUID(Structure):
            _fields_ = [
                ("Data1", DWORD),
                ("Data2", c_ushort),
                ("Data3", c_ushort),
                ("Data4", c_ubyte * 8)
            ]

        IID_IAccessible = GUID(
            0x618736E0, 0x3C3D, 0x11CF,
            (c_ubyte * 8)(0x81, 0x0C, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71)
        )
        IID_NULL = GUID(0, 0, 0, (c_ubyte * 8)(0, 0, 0, 0, 0, 0, 0, 0))

        oleacc = windll.oleacc
        obj_ptr = c_void_p()

        result = oleacc.AccessibleObjectFromWindow(
            hwnd, 0, byref(IID_IAccessible), byref(obj_ptr)
        )

        if result != 0 or not obj_ptr.value:
            return {"success": False, "message": "IAccessible 획득 실패"}

        print(f"✓ IAccessible: 0x{obj_ptr.value:08X}")

        # VARIANT 정의 (64비트 고려)
        class VARIANT(Structure):
            _fields_ = [
                ("vt", c_ushort),
                ("wReserved1", c_ushort),
                ("wReserved2", c_ushort),
                ("wReserved3", c_ushort),
                ("val", c_ulonglong)  # 포인터 또는 값
            ]

        class DISPPARAMS(Structure):
            _fields_ = [
                ("rgvarg", POINTER(VARIANT)),
                ("rgdispidNamedArgs", POINTER(LONG)),
                ("cArgs", c_uint),
                ("cNamedArgs", c_uint)
            ]

        INVOKE_FUNC = WINFUNCTYPE(
            HRESULT, c_void_p, LONG, POINTER(GUID), DWORD, c_ushort,
            POINTER(DISPPARAMS), POINTER(VARIANT), c_void_p, POINTER(c_uint)
        )

        vtable = cast(obj_ptr, POINTER(c_void_p)).contents
        vtable_array = cast(vtable, POINTER(c_void_p))
        invoke_func = INVOKE_FUNC(vtable_array[6])

        DISPATCH_PROPERTYGET = 2
        DISPID_ACC_FOCUS = -5011
        DISPID_ACC_VALUE = -5004

        print("\n=== accFocus 호출 ===")
        params = DISPPARAMS()
        params.cArgs = 0
        result_variant = VARIANT()

        hr = invoke_func(
            obj_ptr, DISPID_ACC_FOCUS, byref(IID_NULL), 0,
            DISPATCH_PROPERTYGET, byref(params), byref(result_variant),
            None, None
        )

        print(f"결과: 0x{hr:08X}, VARIANT 타입: {result_variant.vt}")

        if hr == 0 and result_variant.vt == 9:  # VT_DISPATCH
            focused_ptr = c_void_p(result_variant.val)
            print(f"✓ 포커스된 객체: 0x{focused_ptr.value:08X}")

            # 포커스된 객체의 accValue 호출
            print("\n=== 포커스된 객체의 accValue ===")

            vtable2 = cast(focused_ptr, POINTER(c_void_p)).contents
            vtable_array2 = cast(vtable2, POINTER(c_void_p))
            invoke_func2 = INVOKE_FUNC(vtable_array2[6])

            childid_self = VARIANT()
            childid_self.vt = 3  # VT_I4
            childid_self.val = 0  # CHILDID_SELF

            params2 = DISPPARAMS()
            params2.rgvarg = pointer(childid_self)
            params2.cArgs = 1

            result_variant2 = VARIANT()

            hr2 = invoke_func2(
                focused_ptr, DISPID_ACC_VALUE, byref(IID_NULL), 0,
                DISPATCH_PROPERTYGET, byref(params2), byref(result_variant2),
                None, None
            )

            print(f"결과: 0x{hr2:08X}, VARIANT 타입: {result_variant2.vt}")

            if hr2 == 0:
                if result_variant2.vt == 8:  # VT_BSTR
                    # BSTR은 c_wchar_p와 호환됨
                    bstr_ptr_value = result_variant2.val
                    print(f"BSTR 포인터 값: 0x{bstr_ptr_value:016X}")

                    if bstr_ptr_value:
                        # BSTR을 문자열로 변환
                        # BSTR은 길이가 앞에 있는 특수 문자열 포맷
                        bstr_ptr = c_wchar_p(bstr_ptr_value)
                        value = bstr_ptr.value

                        print(f"✓ 값: '{value}'")

                        if value == reference_value:
                            print("  ✓✓ 참조 값과 일치!")

                            # 백그라운드 테스트
                            print("\n=== 백그라운드 테스트 ===")
                            print("메모장 실행...")
                            notepad = subprocess.Popen(['notepad.exe'])
                            time.sleep(2)

                            import win32gui
                            active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                            print(f"현재 활성 창: '{active}'")

                            # 다시 accValue 호출
                            result_variant3 = VARIANT()
                            hr3 = invoke_func2(
                                focused_ptr, DISPID_ACC_VALUE, byref(IID_NULL), 0,
                                DISPATCH_PROPERTYGET, byref(params2), byref(result_variant3),
                                None, None
                            )

                            if hr3 == 0 and result_variant3.vt == 8:
                                bstr_ptr3 = c_wchar_p(result_variant3.val)
                                if bstr_ptr3.value:
                                    bg_value = bstr_ptr3.value
                                    print(f"✓✓✓ 백그라운드 값: '{bg_value}'")

                                    if bg_value == reference_value:
                                        print("    ✓✓✓✓ 성공! 백그라운드에서도 값 일치!")

                                        notepad.terminate()
                                        capture_func("attempt65_01_success.png")

                                        return {
                                            "success": True,
                                            "message": f"🎉 IAccessible로 백그라운드 셀 값 읽기 성공! 값='{bg_value}'"
                                        }

                            notepad.terminate()
                    else:
                        print("✗ BSTR 포인터가 NULL")
                else:
                    print(f"✗ 예상치 못한 VARIANT 타입: {result_variant2.vt}")

        capture_func("attempt65_01_complete.png")

        return {
            "success": False,
            "message": "셀 값 읽기 실패"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
