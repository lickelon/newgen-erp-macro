"""
시도 103: COM 인터페이스로 스프레드 읽기
"""
import win32process
import win32gui
import win32con
import win32com.client
from pywinauto import Application


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 103: COM 인터페이스로 스프레드 읽기")
    print("="*60)

    try:
        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[1단계] 급여자료입력 연결 (PID: {process_id})")

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
        print("\n[2단계] 스프레드 찾기")

        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        hwnd = left_spread.handle
        print(f"왼쪽 스프레드: 0x{hwnd:08X}")

        # 3. COM 인터페이스 접근 시도
        print("\n[3단계] COM 인터페이스 접근")

        try:
            # 방법 1: hwnd로 직접 COM 객체 얻기
            print("\n[방법 1] HWND로 COM 객체 얻기")
            try:
                from ctypes import pythonapi, c_void_p, py_object, c_long
                import comtypes

                # ObjectFromLresult 시도
                pythonapi.PyCapsule_GetPointer.restype = c_void_p
                pythonapi.PyCapsule_GetPointer.argtypes = [py_object]

                lres = win32gui.SendMessage(hwnd, win32con.WM_GETOBJECT, 0, 0)
                print(f"  WM_GETOBJECT 응답: {lres}")

            except Exception as e:
                print(f"  ✗ 실패: {e}")

            # 방법 2: pywinauto wrapper 사용
            print("\n[방법 2] pywinauto wrapper 속성 확인")
            try:
                wrapper = left_spread.wrapper_object()
                print(f"  Wrapper 타입: {type(wrapper)}")

                # 일반적인 속성들 시도
                test_attrs = ['Text', 'Value', 'CellText', 'Data', 'Col', 'Row',
                              'MaxCols', 'MaxRows', 'ActiveCol', 'ActiveRow']

                for attr in test_attrs:
                    try:
                        if hasattr(wrapper, attr):
                            val = getattr(wrapper, attr)
                            print(f"  ✓ {attr}: {val}")
                    except:
                        pass

            except Exception as e:
                print(f"  ✗ 실패: {e}")

            # 방법 3: SendMessage로 텍스트 길이/내용 확인
            print("\n[방법 3] SendMessage로 텍스트 확인")
            try:
                import ctypes

                # WM_GETTEXT
                buffer_size = 1024
                buffer = ctypes.create_unicode_buffer(buffer_size)
                length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXT,
                                             buffer_size, ctypes.addressof(buffer))

                if length > 0:
                    text = buffer.value
                    print(f"  ✓ WM_GETTEXT: '{text}' (길이: {length})")
                else:
                    print(f"  ✗ WM_GETTEXT 응답 없음")

            except Exception as e:
                print(f"  ✗ 실패: {e}")

            # 방법 4: 커스텀 메시지 시도
            print("\n[방법 4] 커스텀 메시지 시도")
            try:
                # Farpoint Spread의 일반적인 메시지들
                # SSM_GETTEXT = WM_USER + 100 정도로 추정
                base = win32con.WM_USER

                test_messages = [
                    (base + 100, "WM_USER+100"),
                    (base + 101, "WM_USER+101"),
                    (base + 200, "WM_USER+200"),
                ]

                for msg, desc in test_messages:
                    result = win32gui.SendMessage(hwnd, msg, 0, 0)
                    if result != 0:
                        print(f"  ✓ {desc}: {result}")

            except Exception as e:
                print(f"  ✗ 실패: {e}")

        except Exception as e:
            print(f"COM 접근 실패: {e}")

        return {"success": False, "message": "COM 인터페이스로 읽기 실패 - 추가 조사 필요"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
