"""
시도 63: IAccessible ctypes로 직접 호출

comtypes.gen 없이 ctypes로 직접 IAccessible 메서드 호출
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, LPWSTR, DWORD


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 63: IAccessible ctypes로 직접 호출")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt63_00_initial.png")

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

        # IAccessible GUID: {618736E0-3C3D-11CF-810C-00AA00389B71}
        class GUID(Structure):
            _fields_ = [
                ("Data1", DWORD),
                ("Data2", c_ushort),
                ("Data3", c_ushort),
                ("Data4", c_ubyte * 8)
            ]

        IID_IAccessible = GUID(
            0x618736E0,
            0x3C3D,
            0x11CF,
            (c_ubyte * 8)(0x81, 0x0C, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71)
        )

        oleacc = windll.oleacc
        obj_ptr = c_void_p()

        result = oleacc.AccessibleObjectFromWindow(
            hwnd,
            0,  # OBJID_CLIENT
            byref(IID_IAccessible),
            byref(obj_ptr)
        )

        if result != 0 or not obj_ptr.value:
            return {"success": False, "message": f"IAccessible 획득 실패: {result}"}

        print(f"✓ IAccessible 포인터: 0x{obj_ptr.value:08X}")

        # IAccessible vtable 정의 (간소화 버전)
        # IUnknown: QueryInterface, AddRef, Release
        # IDispatch: GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke
        # IAccessible: 22 메서드들...

        # accName 메서드 호출 (IDispatch::Invoke 사용)
        print("\n=== IDispatch::Invoke로 accValue 호출 ===")

        # IDispatch vtable offset
        # 0: QueryInterface
        # 1: AddRef
        # 2: Release
        # 3: GetTypeInfoCount
        # 4: GetTypeInfo
        # 5: GetIDsOfNames
        # 6: Invoke

        # Invoke 함수 프로토타입
        # HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
        #                DISPPARAMS *pDispParams, VARIANT *pVarResult,
        #                EXCEPINFO *pExcepInfo, UINT *puArgErr);

        class VARIANT(Structure):
            _fields_ = [
                ("vt", c_ushort),
                ("wReserved1", c_ushort),
                ("wReserved2", c_ushort),
                ("wReserved3", c_ushort),
                ("union", c_ubyte * 8)  # 간소화
            ]

        class DISPPARAMS(Structure):
            _fields_ = [
                ("rgvarg", POINTER(VARIANT)),
                ("rgdispidNamedArgs", POINTER(LONG)),
                ("cArgs", c_uint),
                ("cNamedArgs", c_uint)
            ]

        # accValue DISPID (보통 -5004)
        DISPID_ACC_VALUE = -5004
        DISPID_ACC_NAME = -5003

        # vtable에서 Invoke 함수 포인터 얻기
        vtable = cast(obj_ptr, POINTER(c_void_p)).contents
        vtable_array = cast(vtable, POINTER(c_void_p))

        invoke_ptr = vtable_array[6]  # IDispatch::Invoke

        # Invoke 함수 타입 정의
        INVOKE_FUNC = WINFUNCTYPE(
            HRESULT,  # 반환값
            c_void_p,  # this
            LONG,  # dispIdMember
            POINTER(GUID),  # riid
            DWORD,  # lcid
            c_ushort,  # wFlags
            POINTER(DISPPARAMS),  # pDispParams
            POINTER(VARIANT),  # pVarResult
            c_void_p,  # pExcepInfo
            POINTER(c_uint)  # puArgErr
        )

        invoke_func = INVOKE_FUNC(invoke_ptr)

        # CHILDID_SELF를 파라미터로 전달
        childid_variant = VARIANT()
        childid_variant.vt = 3  # VT_I4 (정수)
        cast(byref(childid_variant, 8), POINTER(LONG)).contents = LONG(0)  # CHILDID_SELF = 0

        params = DISPPARAMS()
        params.rgvarg = pointer(childid_variant)
        params.cArgs = 1
        params.cNamedArgs = 0

        result_variant = VARIANT()

        IID_NULL = GUID(0, 0, 0, (c_ubyte * 8)(0, 0, 0, 0, 0, 0, 0, 0))

        # accValue 호출
        DISPATCH_PROPERTYGET = 2
        hr = invoke_func(
            obj_ptr,
            DISPID_ACC_VALUE,
            byref(IID_NULL),
            0,  # LOCALE_USER_DEFAULT
            DISPATCH_PROPERTYGET,
            byref(params),
            byref(result_variant),
            None,
            None
        )

        print(f"Invoke 결과: 0x{hr:08X}")

        if hr == 0:  # S_OK
            print(f"VARIANT 타입: {result_variant.vt}")

            if result_variant.vt == 8:  # VT_BSTR (문자열)
                # BSTR 포인터 추출
                bstr_ptr = cast(byref(result_variant, 8), POINTER(c_wchar_p)).contents
                if bstr_ptr:
                    value = bstr_ptr.value
                    print(f"✓ accValue: '{value}'")

                    if value == reference_value:
                        print("  ✓✓ 참조 값과 일치!")

                        # 백그라운드 테스트
                        print("\n=== 백그라운드 테스트 ===")
                        print("메모장 실행...")
                        notepad = subprocess.Popen(['notepad.exe'])
                        time.sleep(2)

                        # 다시 accValue 호출
                        result_variant2 = VARIANT()
                        hr2 = invoke_func(
                            obj_ptr,
                            DISPID_ACC_VALUE,
                            byref(IID_NULL),
                            0,
                            DISPATCH_PROPERTYGET,
                            byref(params),
                            byref(result_variant2),
                            None,
                            None
                        )

                        if hr2 == 0:
                            bstr_ptr2 = cast(byref(result_variant2, 8), POINTER(c_wchar_p)).contents
                            if bstr_ptr2:
                                bg_value = bstr_ptr2.value
                                print(f"✓✓ 백그라운드에서 값 읽기 성공: '{bg_value}'")

                                notepad.terminate()
                                capture_func("attempt63_01_success.png")

                                return {
                                    "success": True,
                                    "message": f"IAccessible로 백그라운드 읽기 성공! 값='{bg_value}'"
                                }

                        notepad.terminate()
                else:
                    print("✗ BSTR 포인터가 NULL")
            else:
                print(f"✗ 예상치 못한 VARIANT 타입: {result_variant.vt}")
        else:
            print(f"✗ Invoke 실패: 0x{hr:08X}")

        capture_func("attempt63_01_complete.png")

        return {
            "success": False,
            "message": "IAccessible로 값 읽기 실패"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
