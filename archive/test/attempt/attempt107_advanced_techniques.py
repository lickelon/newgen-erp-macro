"""
시도 107: 고급 기법들 시도

- UI Automation 패턴
- 윈도우 속성 깊이 분석
- 툴팁 확인
- 메모리 읽기 시도
"""
import win32process
import win32gui
import win32con
import win32api
import time
import pyperclip
from pywinauto import Application
from ctypes import create_unicode_buffer, windll, c_long, byref, sizeof
import struct


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 107: 고급 기법들 시도")
    print("="*60)

    try:
        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[단계] 급여자료입력 연결 (PID: {process_id})")

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
        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        left_hwnd = left_spread.handle
        print(f"\n왼쪽 스프레드: 0x{left_hwnd:08X}")

        # === 방법 1: UI Automation 백엔드 시도 ===
        print("\n" + "="*60)
        print("[방법 1] UI Automation 백엔드 시도")
        print("="*60)

        try:
            from pywinauto import Application as UIA_App
            uia_app = UIA_App(backend="uia")
            uia_app.connect(title="급여자료입력")

            print("✓ UIA 연결 성공")

            # UIA로 스프레드 찾기
            try:
                uia_dlg = uia_app.window(title="급여자료입력")
                print("✓ UIA 윈도우 찾기 성공")

                # 모든 descendants 확인
                all_controls = uia_dlg.descendants()
                print(f"✓ UIA descendants: {len(all_controls)}개")

                # 데이터 그리드나 테이블 찾기
                for ctrl in all_controls[:50]:  # 처음 50개만
                    try:
                        ctrl_type = ctrl.element_info.control_type
                        name = ctrl.element_info.name
                        if ctrl_type in ["DataGrid", "Table", "List", "DataItem"]:
                            print(f"  [{ctrl_type}] '{name}'")

                            # 텍스트 읽기 시도
                            try:
                                text = ctrl.window_text()
                                if text:
                                    print(f"    텍스트: '{text}'")
                            except:
                                pass
                    except:
                        pass

            except Exception as e:
                print(f"✗ UIA 윈도우 찾기 실패: {e}")

        except Exception as e:
            print(f"✗ UIA 백엔드 실패: {e}")

        # === 방법 2: GetWindowLongPtr로 추가 정보 ===
        print("\n" + "="*60)
        print("[방법 2] GetWindowLongPtr로 윈도우 데이터 확인")
        print("="*60)

        GWL_USERDATA = -21
        GWLP_USERDATA = -21

        try:
            user_data = win32api.GetWindowLong(left_hwnd, GWL_USERDATA)
            print(f"GWL_USERDATA: 0x{user_data:X}")

            if user_data != 0:
                print(f"  → 사용자 데이터 포인터 발견!")
        except Exception as e:
            print(f"✗ GetWindowLong 실패: {e}")

        # === 방법 3: 툴팁 확인 ===
        print("\n" + "="*60)
        print("[방법 3] 셀에 마우스 올렸을 때 툴팁 확인")
        print("="*60)

        # 셀 위치에 마우스 올리기 (WM_MOUSEMOVE)
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
        time.sleep(0.5)

        # 툴팁 윈도우 찾기
        tooltip_found = False
        for ctrl in installment_dlg.descendants():
            try:
                cn = ctrl.class_name()
                if "tooltip" in cn.lower() or cn == "tooltips_class32":
                    text = ctrl.window_text()
                    if text:
                        print(f"✓✓✓ 툴팁 발견: '{text}'")
                        tooltip_found = True
                        return {"success": True, "message": f"툴팁에서 읽기 성공: '{text}'"}
            except:
                pass

        if not tooltip_found:
            print("✗ 툴팁 없음")

        # === 방법 4: WM_GETTEXT 큰 버퍼로 시도 ===
        print("\n" + "="*60)
        print("[방법 4] WM_GETTEXT 큰 버퍼 (10KB)")
        print("="*60)

        try:
            buffer_size = 10240  # 10KB
            buffer = create_unicode_buffer(buffer_size)
            length = win32gui.SendMessage(left_hwnd, win32con.WM_GETTEXT, buffer_size, buffer)

            if length > 0:
                text = buffer.value
                print(f"✓✓✓ WM_GETTEXT 성공: {length}바이트")
                print(f"내용: '{text[:200]}'...")
                return {"success": True, "message": f"WM_GETTEXT로 읽기 성공: {length}바이트"}
            else:
                print(f"✗ WM_GETTEXT: 0바이트")
        except Exception as e:
            print(f"✗ WM_GETTEXT 실패: {e}")

        # === 방법 5: 스프레드의 모든 자식 윈도우 열거 ===
        print("\n" + "="*60)
        print("[방법 5] EnumChildWindows로 깊은 탐색")
        print("="*60)

        child_windows = []

        def enum_child_proc(hwnd, lParam):
            try:
                cn = win32gui.GetClassName(hwnd)
                text = win32gui.GetWindowText(hwnd)
                child_windows.append((hwnd, cn, text))
            except:
                pass
            return True

        try:
            win32gui.EnumChildWindows(left_hwnd, enum_child_proc, None)

            if child_windows:
                print(f"✓ 자식 윈도우 {len(child_windows)}개 발견:")
                for hwnd, cn, text in child_windows[:10]:
                    if text:
                        print(f"  [0x{hwnd:08X}] [{cn}] '{text}'")

                # 텍스트가 있는 첫 번째 자식
                for hwnd, cn, text in child_windows:
                    if text and len(text) > 0:
                        print(f"\n✓✓✓ 자식에서 텍스트 발견: '{text}'")
                        return {"success": True, "message": f"자식 윈도우에서 읽기: '{text}'"}
            else:
                print("✗ 자식 윈도우 없음")
        except Exception as e:
            print(f"✗ EnumChildWindows 실패: {e}")

        # === 방법 6: GetWindowInfo 구조체 확인 ===
        print("\n" + "="*60)
        print("[방법 6] GetWindowInfo 구조체 분석")
        print("="*60)

        try:
            from ctypes.wintypes import DWORD, RECT, ATOM

            class WINDOWINFO(struct.Structure):
                _fields_ = [
                    ("cbSize", DWORD),
                    ("rcWindow", RECT),
                    ("rcClient", RECT),
                    ("dwStyle", DWORD),
                    ("dwExStyle", DWORD),
                    ("dwWindowStatus", DWORD),
                    ("cxWindowBorders", c_long),
                    ("cyWindowBorders", c_long),
                    ("atomWindowType", ATOM),
                    ("wCreatorVersion", c_long),
                ]

            info = WINDOWINFO()
            info.cbSize = sizeof(WINDOWINFO)

            result = windll.user32.GetWindowInfo(left_hwnd, byref(info))
            if result:
                print(f"✓ GetWindowInfo 성공")
                print(f"  Style: 0x{info.dwStyle:08X}")
                print(f"  ExStyle: 0x{info.dwExStyle:08X}")
                print(f"  Status: 0x{info.dwWindowStatus:08X}")
                print(f"  Type: 0x{info.atomWindowType:04X}")
            else:
                print("✗ GetWindowInfo 실패")
        except Exception as e:
            print(f"✗ GetWindowInfo 오류: {e}")

        # === 방법 7: Shift+방향키로 범위 선택 ===
        print("\n" + "="*60)
        print("[방법 7] Shift+방향키로 범위 선택 후 복사")
        print("="*60)

        # Home 키
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.2)

        # Shift+End (행 전체 선택)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_SHIFT, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_END, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_END, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_SHIFT, 0)
        time.sleep(0.3)

        # 복사
        pyperclip.copy("")
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        value = pyperclip.paste()
        if value:
            print(f"✓✓✓ Shift+End 후 복사 성공: '{value}'")
            return {"success": True, "message": f"Shift+End 후 복사 성공: '{value}'"}
        else:
            print("✗ Shift+End 후에도 복사 안 됨")

        # === 방법 8: 스프레드 속성 페이지 열기 (디버그) ===
        print("\n" + "="*60)
        print("[방법 8] Alt+Enter로 속성 확인")
        print("="*60)

        # Alt+Enter
        win32api.SendMessage(left_hwnd, win32con.WM_SYSKEYDOWN, win32con.VK_RETURN, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_SYSKEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.5)

        # 새 윈도우가 열렸는지 확인
        print("✓ Alt+Enter 전송 완료")

        return {
            "success": False,
            "message": "8가지 고급 기법 모두 실패"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
