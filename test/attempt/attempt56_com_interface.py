"""
시도 56: COM 인터페이스로 fpUSpread80 셀 값 직접 읽기

Farpoint Spread는 ActiveX 컨트롤이므로 COM 인터페이스를 통해
백그라운드에서도 셀 값을 읽을 수 있을 것으로 예상
"""
import time
import win32com.client
import pythoncom


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 56: COM 인터페이스로 셀 값 읽기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt56_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")
        print(f"위치: {left_spread.rectangle()}")

        print("\n=== 방법 1: win32com.client로 HWND 연결 시도 ===")
        try:
            # HWND로부터 COM 객체 가져오기 시도
            pythoncom.CoInitialize()

            # AccessibleObjectFromWindow를 통한 접근 시도
            from ctypes import windll, byref, c_void_p
            from comtypes import GUID

            IID_IDispatch = GUID("{00020400-0000-0000-C000-000000000046}")
            ptr = c_void_p()

            # oleacc.dll 사용
            oleacc = windll.LoadLibrary("oleacc.dll")
            result = oleacc.AccessibleObjectFromWindow(
                hwnd,
                0,  # OBJID_CLIENT
                byref(IID_IDispatch),
                byref(ptr)
            )

            if result == 0 and ptr:
                print("✓ COM 객체 획득 성공")
                com_obj = win32com.client.Dispatch(ptr)
                print(f"  타입: {type(com_obj)}")

                # 메서드 탐색
                print("\n사용 가능한 메서드/속성:")
                for attr in dir(com_obj):
                    if not attr.startswith('_'):
                        print(f"  - {attr}")

            else:
                print(f"✗ AccessibleObjectFromWindow 실패 (result={result})")

        except Exception as e:
            print(f"✗ 방법 1 실패: {e}")

        print("\n=== 방법 2: pywinauto의 COM 속성 확인 ===")
        try:
            # pywinauto wrapper가 제공하는 속성 확인
            print("left_spread 속성:")
            for attr in dir(left_spread):
                if not attr.startswith('_') and 'com' in attr.lower():
                    print(f"  - {attr}: {getattr(left_spread, attr, 'N/A')}")

        except Exception as e:
            print(f"✗ 방법 2 실패: {e}")

        print("\n=== 방법 3: GetText 등 표준 메서드 시도 ===")
        try:
            # 일반적인 Spread 메서드들 시도
            methods_to_try = [
                ('GetText', [0, 0]),  # GetText(col, row)
                ('GetText', [1, 1]),
                ('get_Text', [0, 0]),
                ('Text', None),
                ('GetValue', [0, 0]),
                ('Value', None),
                ('GetData', [0, 0]),
                ('Data', None),
            ]

            for method_name, args in methods_to_try:
                try:
                    if hasattr(left_spread, method_name):
                        method = getattr(left_spread, method_name)
                        if args:
                            result = method(*args)
                        else:
                            result = method
                        print(f"✓ {method_name}: {result}")
                    else:
                        print(f"✗ {method_name}: 메서드 없음")
                except Exception as e:
                    print(f"✗ {method_name}: {e}")

        except Exception as e:
            print(f"✗ 방법 3 실패: {e}")

        print("\n=== 방법 4: SendMessage로 텍스트 길이/텍스트 가져오기 ===")
        try:
            import win32api
            import win32con
            from ctypes import create_string_buffer, c_int, sizeof

            # WM_GETTEXT 시도
            WM_GETTEXT = 0x000D
            WM_GETTEXTLENGTH = 0x000E

            # 텍스트 길이 확인
            length = win32api.SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
            print(f"텍스트 길이: {length}")

            if length > 0:
                # 텍스트 가져오기
                buffer = create_string_buffer(length + 1)
                win32api.SendMessage(hwnd, WM_GETTEXT, length + 1, buffer)
                text = buffer.value.decode('utf-8', errors='ignore')
                print(f"✓ WM_GETTEXT: '{text}'")
            else:
                print("✗ WM_GETTEXT: 텍스트 길이 0")

        except Exception as e:
            print(f"✗ 방법 4 실패: {e}")

        print("\n=== 방법 5: Spy++ 스타일 메시지 조사 ===")
        try:
            # 사용자 정의 메시지 범위 조사
            print("사용자 정의 메시지 시도 중...")

            # Farpoint Spread의 일반적인 메시지들
            # (문서나 헤더 파일이 필요하지만, 일반적인 값들 시도)
            test_messages = [
                (0x8000 + 100, "Custom+100"),
                (0x8000 + 200, "Custom+200"),
                (0x400, "WM_USER"),
                (0x400 + 100, "WM_USER+100"),
            ]

            for msg, name in test_messages:
                try:
                    result = win32api.SendMessage(hwnd, msg, 0, 0)
                    if result != 0:
                        print(f"  메시지 {name} (0x{msg:04X}): {result}")
                except:
                    pass

        except Exception as e:
            print(f"✗ 방법 5 실패: {e}")

        capture_func("attempt56_01_complete.png")

        return {
            "success": False,
            "message": "COM 인터페이스 직접 접근 실패 - 대안 필요"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
