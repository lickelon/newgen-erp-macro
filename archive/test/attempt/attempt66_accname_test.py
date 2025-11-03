"""
시도 66: accName, accDescription 등 모든 속성 테스트

accValue가 NULL이므로 다른 IAccessible 속성들로 값 읽기 시도
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, DWORD


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 66: accName 등 모든 속성 테스트")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt66_00_initial.png")

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

        # VARIANT 정의
        class VARIANT(Structure):
            _fields_ = [
                ("vt", c_ushort),
                ("wReserved1", c_ushort),
                ("wReserved2", c_ushort),
                ("wReserved3", c_ushort),
                ("val", c_ulonglong)
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

        print("\n=== accFocus 호출 ===")
        params = DISPPARAMS()
        params.cArgs = 0
        result_variant = VARIANT()

        hr = invoke_func(
            obj_ptr, -5011, byref(IID_NULL), 0,
            DISPATCH_PROPERTYGET, byref(params), byref(result_variant),
            None, None
        )

        if hr != 0 or result_variant.vt != 9:
            return {"success": False, "message": "accFocus 실패"}

        focused_ptr = c_void_p(result_variant.val)
        print(f"✓ 포커스된 객체: 0x{focused_ptr.value:08X}")

        vtable2 = cast(focused_ptr, POINTER(c_void_p)).contents
        vtable_array2 = cast(vtable2, POINTER(c_void_p))
        invoke_func2 = INVOKE_FUNC(vtable_array2[6])

        childid_self = VARIANT()
        childid_self.vt = 3
        childid_self.val = 0

        params2 = DISPPARAMS()
        params2.rgvarg = pointer(childid_self)
        params2.cArgs = 1

        # 모든 IAccessible 속성 테스트
        properties = [
            (-5003, "accName"),
            (-5004, "accValue"),
            (-5005, "accDescription"),
            (-5009, "accHelp"),
            (-5013, "accKeyboardShortcut"),
            (-5014, "accDefaultAction"),
        ]

        print("\n=== 포커스된 객체의 모든 속성 ===")
        found_value = None

        for dispid, prop_name in properties:
            result_prop = VARIANT()
            hr = invoke_func2(
                focused_ptr, dispid, byref(IID_NULL), 0,
                DISPATCH_PROPERTYGET, byref(params2), byref(result_prop),
                None, None
            )

            print(f"{prop_name} (DISPID={dispid}):")
            print(f"  결과: 0x{hr:08X}, VARIANT 타입: {result_prop.vt}")

            if hr == 0 and result_prop.vt == 8:  # VT_BSTR
                if result_prop.val:
                    bstr_ptr = c_wchar_p(result_prop.val)
                    value = bstr_ptr.value
                    print(f"  ✓ 값: '{value}'")

                    if value == reference_value:
                        print(f"    ✓✓ 참조 값과 일치!")
                        found_value = (prop_name, value, dispid)
                else:
                    print(f"  ✗ BSTR 포인터 NULL")
            elif hr == 0 and result_prop.vt == 0:  # VT_EMPTY
                print(f"  (비어있음)")
            elif hr == 0:
                print(f"  (다른 타입: {result_prop.vt})")

        if found_value:
            prop_name, value, dispid = found_value
            print(f"\n✓✓✓ 성공! {prop_name}에서 값 발견: '{value}'")

            # 백그라운드 테스트
            print("\n=== 백그라운드 테스트 ===")
            print("메모장 실행...")
            notepad = subprocess.Popen(['notepad.exe'])
            time.sleep(2)

            import win32gui
            active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            print(f"현재 활성 창: '{active}'")

            # 다시 같은 속성 호출
            result_bg = VARIANT()
            hr_bg = invoke_func2(
                focused_ptr, dispid, byref(IID_NULL), 0,
                DISPATCH_PROPERTYGET, byref(params2), byref(result_bg),
                None, None
            )

            if hr_bg == 0 and result_bg.vt == 8 and result_bg.val:
                bstr_bg = c_wchar_p(result_bg.val)
                bg_value = bstr_bg.value
                print(f"✓✓✓ 백그라운드 값: '{bg_value}'")

                if bg_value == reference_value:
                    print("    ✓✓✓✓ 성공! 백그라운드에서도 값 일치!")

                    notepad.terminate()
                    capture_func("attempt66_01_success.png")

                    return {
                        "success": True,
                        "message": f"🎉 {prop_name}으로 백그라운드 셀 값 읽기 성공! 값='{bg_value}'"
                    }

            notepad.terminate()

        capture_func("attempt66_01_complete.png")

        return {
            "success": False,
            "message": "모든 IAccessible 속성에서 값 읽기 실패"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
