"""
시도 62: IAccessible로 셀 값 읽기

AccessibleObjectFromWindow로 얻은 IAccessible 인터페이스를 사용하여
accValue, accName 등으로 현재 셀의 값 읽기
"""
import time
import pythoncom
from ctypes import *
from comtypes import GUID, IUnknown, POINTER, byref
import comtypes.client
import comtypes.automation
import subprocess


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 62: IAccessible로 셀 값 읽기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt62_00_initial.png")

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
        pythoncom.CoInitialize()

        # IAccessible GUID
        IID_IAccessible = GUID("{618736E0-3C3D-11CF-810C-00AA00389B71}")

        oleacc = windll.oleacc
        obj_ptr = POINTER(IUnknown)()

        result = oleacc.AccessibleObjectFromWindow(
            hwnd,
            0,  # OBJID_CLIENT
            byref(IID_IAccessible),
            byref(obj_ptr)
        )

        if result != 0 or not obj_ptr:
            return {"success": False, "message": f"IAccessible 획득 실패: {result}"}

        print(f"✓ IAccessible 포인터 획득")

        # IAccessible로 캐스팅
        from comtypes.gen.Accessibility import IAccessible as IAccessibleInterface
        acc = obj_ptr.QueryInterface(IAccessibleInterface)
        print(f"✓ IAccessible 인터페이스: {acc}")

        print("\n=== 현재 객체 정보 ===")
        # CHILDID_SELF = 0
        CHILDID_SELF = 0

        try:
            name = acc.accName(CHILDID_SELF)
            print(f"이름: '{name}'")
        except:
            print("이름: (없음)")

        try:
            value = acc.accValue(CHILDID_SELF)
            print(f"값: '{value}'")
            if value == reference_value:
                print("  ✓✓ 참조 값과 일치!")
        except:
            print("값: (없음)")

        try:
            role = acc.accRole(CHILDID_SELF)
            print(f"역할: {role}")
        except:
            print("역할: (없음)")

        try:
            state = acc.accState(CHILDID_SELF)
            print(f"상태: {state}")
        except:
            print("상태: (없음)")

        print("\n=== 자식 객체 탐색 ===")
        try:
            child_count = acc.accChildCount
            print(f"자식 개수: {child_count}")

            if child_count > 0:
                # 모든 자식 열거
                from ctypes import c_long
                from comtypes.automation import VARIANT

                # accChildren 배열 얻기
                children_array = (VARIANT * child_count)()
                obtained = c_long()

                result = oleacc.AccessibleChildren(
                    acc,
                    0,  # 시작 인덱스
                    child_count,
                    children_array,
                    byref(obtained)
                )

                if result == 0:
                    print(f"✓ {obtained.value}개 자식 획득")

                    for i in range(obtained.value):
                        child_var = children_array[i]

                        # VARIANT 타입 확인
                        if child_var.vt == 3:  # VT_I4 (정수 = child ID)
                            child_id = child_var.value

                            try:
                                child_name = acc.accName(child_id)
                            except:
                                child_name = None

                            try:
                                child_value = acc.accValue(child_id)
                            except:
                                child_value = None

                            if child_name or child_value:
                                print(f"  자식 {child_id}:")
                                if child_name:
                                    print(f"    이름: '{child_name}'")
                                if child_value:
                                    print(f"    값: '{child_value}'")
                                    if child_value == reference_value or reference_value in str(child_value):
                                        print(f"      ✓✓ 참조 값 발견!")

                        elif child_var.vt == 9:  # VT_DISPATCH (IDispatch 포인터)
                            print(f"  자식 {i}: IDispatch 객체")
                            # IDispatch를 IAccessible로 변환 가능

        except Exception as e:
            import traceback
            print(f"자식 탐색 실패: {e}")
            traceback.print_exc()

        print("\n=== 백그라운드 테스트 ===")
        try:
            print("메모장 실행하여 창 비활성화...")
            notepad = subprocess.Popen(['notepad.exe'])
            time.sleep(2)

            import win32gui
            active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            print(f"현재 활성 창: '{active_title}'")

            # 백그라운드에서 값 읽기 시도
            try:
                bg_value = acc.accValue(CHILDID_SELF)
                print(f"✓ 백그라운드 값 읽기: '{bg_value}'")

                if bg_value == reference_value:
                    print("  ✓✓ 백그라운드에서도 값 일치!")
                    notepad.terminate()
                    capture_func("attempt62_01_success.png")

                    return {
                        "success": True,
                        "message": f"IAccessible로 백그라운드 읽기 성공! 값='{bg_value}'"
                    }
            except Exception as e:
                print(f"✗ 백그라운드 읽기 실패: {e}")

            notepad.terminate()
            time.sleep(0.5)

        except Exception as e:
            print(f"백그라운드 테스트 실패: {e}")

        capture_func("attempt62_01_complete.png")

        return {
            "success": False,
            "message": "IAccessible로 셀 값 직접 읽기 불가능"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
