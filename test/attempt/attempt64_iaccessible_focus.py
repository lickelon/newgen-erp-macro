"""
시도 64: IAccessible accFocus로 현재 포커스된 셀 찾기

fpUSpread80의 accFocus 또는 accSelection을 사용하여
현재 선택/포커스된 셀의 IAccessible을 얻고 값 읽기
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, DWORD


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 64: IAccessible accFocus로 현재 셀 찾기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt64_00_initial.png")

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

        print("\n=== IAccessible 인터페이스 획득 ===")

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
            return {"success": False, "message": f"IAccessible 획득 실패: {result}"}

        print(f"✓ IAccessible 포인터: 0x{obj_ptr.value:08X}")

        # VARIANT 정의
        class VARIANT(Structure):
            _fields_ = [
                ("vt", c_ushort),
                ("wReserved1", c_ushort),
                ("wReserved2", c_ushort),
                ("wReserved3", c_ushort),
                ("data", c_ulonglong)  # union을 간소화
            ]

        class DISPPARAMS(Structure):
            _fields_ = [
                ("rgvarg", POINTER(VARIANT)),
                ("rgdispidNamedArgs", POINTER(LONG)),
                ("cArgs", c_uint),
                ("cNamedArgs", c_uint)
            ]

        # Invoke 함수 타입
        INVOKE_FUNC = WINFUNCTYPE(
            HRESULT, c_void_p, LONG, POINTER(GUID), DWORD, c_ushort,
            POINTER(DISPPARAMS), POINTER(VARIANT), c_void_p, POINTER(c_uint)
        )

        vtable = cast(obj_ptr, POINTER(c_void_p)).contents
        vtable_array = cast(vtable, POINTER(c_void_p))
        invoke_func = INVOKE_FUNC(vtable_array[6])

        DISPATCH_PROPERTYGET = 2
        DISPATCH_METHOD = 1

        # accFocus DISPID
        DISPID_ACC_FOCUS = -5011
        DISPID_ACC_SELECTION = -5012
        DISPID_ACC_VALUE = -5004
        DISPID_ACC_NAME = -5003

        print("\n=== accFocus 호출 ===")

        # CHILDID_SELF 파라미터
        params = DISPPARAMS()
        params.cArgs = 0

        result_variant = VARIANT()

        hr = invoke_func(
            obj_ptr, DISPID_ACC_FOCUS, byref(IID_NULL), 0,
            DISPATCH_PROPERTYGET, byref(params), byref(result_variant),
            None, None
        )

        print(f"accFocus 결과: 0x{hr:08X}")
        print(f"VARIANT 타입: {result_variant.vt}")

        focused_obj = None

        if hr == 0:
            if result_variant.vt == 3:  # VT_I4 (child ID)
                child_id = cast(byref(result_variant, 8), POINTER(LONG)).contents.value
                print(f"  포커스 자식 ID: {child_id}")

                # 이 child ID로 accValue 호출
                print(f"\n=== 자식 {child_id}의 accValue 호출 ===")

                childid_variant = VARIANT()
                childid_variant.vt = 3
                cast(byref(childid_variant, 8), POINTER(LONG)).contents = LONG(child_id)

                params2 = DISPPARAMS()
                params2.rgvarg = pointer(childid_variant)
                params2.cArgs = 1

                result_variant2 = VARIANT()

                hr2 = invoke_func(
                    obj_ptr, DISPID_ACC_VALUE, byref(IID_NULL), 0,
                    DISPATCH_PROPERTYGET, byref(params2), byref(result_variant2),
                    None, None
                )

                print(f"accValue 결과: 0x{hr2:08X}")

                if hr2 == 0 and result_variant2.vt == 8:  # VT_BSTR
                    bstr_ptr = cast(byref(result_variant2, 8), POINTER(c_wchar_p)).contents
                    if bstr_ptr:
                        value = bstr_ptr.value
                        print(f"✓ 자식의 accValue: '{value}'")

                        if value == reference_value:
                            print("  ✓✓ 참조 값과 일치!")

            elif result_variant.vt == 9:  # VT_DISPATCH (IDispatch 포인터)
                dispatch_ptr = cast(byref(result_variant, 8), POINTER(c_void_p)).contents
                print(f"  포커스 IDispatch: 0x{dispatch_ptr.value:08X}")

                focused_obj = dispatch_ptr

                # 이 IDispatch로 accValue 호출
                print(f"\n=== 포커스된 객체의 accValue 호출 ===")

                vtable2 = cast(dispatch_ptr, POINTER(c_void_p)).contents
                vtable_array2 = cast(vtable2, POINTER(c_void_p))
                invoke_func2 = INVOKE_FUNC(vtable_array2[6])

                childid_self = VARIANT()
                childid_self.vt = 3
                cast(byref(childid_self, 8), POINTER(LONG)).contents = LONG(0)

                params3 = DISPPARAMS()
                params3.rgvarg = pointer(childid_self)
                params3.cArgs = 1

                result_variant3 = VARIANT()

                hr3 = invoke_func2(
                    dispatch_ptr, DISPID_ACC_VALUE, byref(IID_NULL), 0,
                    DISPATCH_PROPERTYGET, byref(params3), byref(result_variant3),
                    None, None
                )

                print(f"accValue 결과: 0x{hr3:08X}")
                print(f"VARIANT 타입: {result_variant3.vt}")

                if hr3 == 0 and result_variant3.vt == 8:  # VT_BSTR
                    bstr_ptr = cast(byref(result_variant3, 8), POINTER(c_wchar_p)).contents
                    if bstr_ptr:
                        value = bstr_ptr.value
                        print(f"✓ 포커스된 객체의 accValue: '{value}'")

                        if value == reference_value:
                            print("  ✓✓ 참조 값과 일치!")

                            # 백그라운드 테스트
                            print("\n=== 백그라운드 테스트 ===")
                            print("메모장 실행...")
                            notepad = subprocess.Popen(['notepad.exe'])
                            time.sleep(2)

                            # 다시 accValue 호출
                            result_variant4 = VARIANT()
                            hr4 = invoke_func2(
                                dispatch_ptr, DISPID_ACC_VALUE, byref(IID_NULL), 0,
                                DISPATCH_PROPERTYGET, byref(params3), byref(result_variant4),
                                None, None
                            )

                            if hr4 == 0 and result_variant4.vt == 8:
                                bstr_ptr2 = cast(byref(result_variant4, 8), POINTER(c_wchar_p)).contents
                                if bstr_ptr2:
                                    bg_value = bstr_ptr2.value
                                    print(f"✓✓✓ 백그라운드에서 값 읽기 성공: '{bg_value}'")

                                    notepad.terminate()
                                    capture_func("attempt64_01_success.png")

                                    return {
                                        "success": True,
                                        "message": f"IAccessible accFocus로 백그라운드 읽기 성공! 값='{bg_value}'"
                                    }

                            notepad.terminate()

        print("\n=== accSelection 시도 ===")
        result_variant5 = VARIANT()

        hr5 = invoke_func(
            obj_ptr, DISPID_ACC_SELECTION, byref(IID_NULL), 0,
            DISPATCH_PROPERTYGET, byref(params), byref(result_variant5),
            None, None
        )

        print(f"accSelection 결과: 0x{hr5:08X}")
        print(f"VARIANT 타입: {result_variant5.vt}")

        capture_func("attempt64_01_complete.png")

        return {
            "success": False,
            "message": "accFocus/accSelection으로 셀 값 읽기 실패"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
