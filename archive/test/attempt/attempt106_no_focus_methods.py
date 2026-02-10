"""
시도 106: 포커스 설정 없이 다양한 방법 시도

set_focus() 사용하지 않고 SendMessage/PostMessage로 직접 제어
"""
import win32process
import win32gui
import win32con
import win32api
import time
import pyperclip
from pywinauto import Application
from ctypes import create_unicode_buffer, windll


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 106: 포커스 설정 없이 다양한 방법 시도")
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
        dlg_hwnd = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            dlg_hwnd = hwnd
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
        right_spread = spreads[1] if len(spreads) > 1 else None

        left_hwnd = left_spread.handle
        print(f"\n왼쪽 스프레드: 0x{left_hwnd:08X}")
        if right_spread:
            right_hwnd = right_spread.handle
            print(f"오른쪽 스프레드: 0x{right_hwnd:08X}")

        # === 방법 1: WM_SETFOCUS로 포커스 설정 ===
        print("\n" + "="*60)
        print("[방법 1] WM_SETFOCUS로 포커스 설정 후 복사")
        print("="*60)

        win32api.SendMessage(left_hwnd, win32con.WM_SETFOCUS, 0, 0)
        time.sleep(0.3)

        pyperclip.copy("")
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        value = pyperclip.paste()
        if value:
            print(f"✓✓✓ WM_SETFOCUS 후 복사 성공: '{value}'")
            return {"success": True, "message": f"WM_SETFOCUS 후 복사 성공: '{value}'"}
        else:
            print("✗ WM_SETFOCUS 후에도 복사 안 됨")

        # === 방법 2: 스프레드 전체 선택 (Ctrl+A) 후 복사 ===
        print("\n" + "="*60)
        print("[방법 2] Ctrl+A 전체 선택 후 복사")
        print("="*60)

        # Ctrl+A
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('A'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('A'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        # Ctrl+C
        pyperclip.copy("")
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        value = pyperclip.paste()
        if value:
            print(f"✓✓✓ Ctrl+A 후 복사 성공: '{value}'")
            return {"success": True, "message": f"Ctrl+A 후 복사 성공: '{value}'"}
        else:
            print("✗ Ctrl+A 후에도 복사 안 됨")

        # === 방법 3: 더블클릭 (좌표 직접 전송) ===
        print("\n" + "="*60)
        print("[방법 3] 더블클릭으로 편집 모드")
        print("="*60)

        # 더블클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lParam)
        time.sleep(0.5)

        # Edit 컨트롤 확인
        edit_found = False
        for ctrl in installment_dlg.descendants():
            try:
                if ctrl.class_name() == "Edit":
                    length = win32gui.SendMessage(ctrl.handle, win32con.WM_GETTEXTLENGTH, 0, 0)
                    if length > 0:
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(ctrl.handle, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value
                        if text:
                            print(f"✓✓✓ Edit 컨트롤 발견: '{text}'")
                            edit_found = True
                            return {"success": True, "message": f"더블클릭 후 Edit에서 읽기: '{text}'"}
            except:
                pass

        if not edit_found:
            print("✗ 더블클릭 후에도 Edit 컨트롤 없음")

        # === 방법 4: Enter 키로 셀 활성화 ===
        print("\n" + "="*60)
        print("[방법 4] Enter 키로 셀 활성화")
        print("="*60)

        # Home 키
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.2)

        # Enter 키
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.3)

        # Edit 컨트롤 확인
        for ctrl in installment_dlg.descendants():
            try:
                if ctrl.class_name() == "Edit":
                    length = win32gui.SendMessage(ctrl.handle, win32con.WM_GETTEXTLENGTH, 0, 0)
                    if length > 0:
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(ctrl.handle, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value
                        if text:
                            print(f"✓✓✓ Enter 후 Edit: '{text}'")
                            return {"success": True, "message": f"Enter 키 후 읽기 성공: '{text}'"}
            except:
                pass

        print("✗ Enter 후에도 Edit 컨트롤 없음")

        # === 방법 5: Space 키 시도 ===
        print("\n" + "="*60)
        print("[방법 5] Space 키로 셀 선택")
        print("="*60)

        # Space 키
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_SPACE, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_SPACE, 0)
        time.sleep(0.3)

        # 복사 시도
        pyperclip.copy("")
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('C'), 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        time.sleep(0.3)

        value = pyperclip.paste()
        if value:
            print(f"✓✓✓ Space 키 후 복사 성공: '{value}'")
            return {"success": True, "message": f"Space 키 후 복사 성공: '{value}'"}
        else:
            print("✗ Space 키 후에도 복사 안 됨")

        # === 방법 6: 행 선택 메시지 시도 ===
        print("\n" + "="*60)
        print("[방법 6] 리스트뷰/리스트박스 메시지 시도")
        print("="*60)

        # LVM_GETITEMTEXT (리스트뷰)
        LVM_FIRST = 0x1000
        LVM_GETITEMTEXT = LVM_FIRST + 45

        # 시도
        result = win32api.SendMessage(left_hwnd, LVM_GETITEMTEXT, 0, 0)
        if result != 0:
            print(f"✓ LVM_GETITEMTEXT 응답: {result}")
        else:
            print("✗ LVM_GETITEMTEXT: 응답 없음 (리스트뷰 아님)")

        # === 방법 7: 컨텍스트 메뉴 열기 (우클릭) ===
        print("\n" + "="*60)
        print("[방법 7] 우클릭 후 복사 메뉴 확인")
        print("="*60)

        # 우클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lParam)
        time.sleep(0.1)
        win32api.SendMessage(left_hwnd, win32con.WM_RBUTTONUP, 0, lParam)
        time.sleep(0.5)

        # 팝업 메뉴가 생성되었는지 확인
        # (실제로는 메뉴 항목을 클릭해야 하지만, 일단 확인만)
        print("✓ 우클릭 완료 (팝업 메뉴 확인 필요)")

        # ESC로 메뉴 닫기
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
        time.sleep(0.2)

        # === 방법 8: 행 전체 상태 메시지 ===
        print("\n" + "="*60)
        print("[방법 8] 커스텀 Spread 메시지 시도")
        print("="*60)

        # Farpoint Spread 커스텀 메시지 (추정)
        base = win32con.WM_USER
        test_messages = [
            (base + 1, "WM_USER+1"),
            (base + 2, "WM_USER+2"),
            (base + 10, "WM_USER+10"),
            (base + 100, "WM_USER+100"),
            (base + 1000, "WM_USER+1000"),
        ]

        for msg, desc in test_messages:
            result = win32api.SendMessage(left_hwnd, msg, 0, 0)
            if result != 0:
                print(f"  {desc}: 0x{result:X}")

        # === 방법 9: 프린트 스크린 방식 (다른 컨트롤에 출력) ===
        print("\n" + "="*60)
        print("[방법 9] 스프레드 데이터가 다른 곳에 표시되는지 확인")
        print("="*60)

        # 첫 행 클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.1)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
        time.sleep(0.5)

        # 모든 Static/Edit에서 값 찾기
        print("클릭 후 Static/Edit 컨트롤 확인:")
        found_texts = []
        for ctrl in installment_dlg.descendants():
            try:
                cn = ctrl.class_name()
                if cn in ["Edit", "Static"]:
                    hwnd = ctrl.handle
                    length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
                    if length > 0 and length < 100:  # 너무 긴 텍스트 제외
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value
                        if text and text not in found_texts:
                            found_texts.append(text)
                            print(f"  [{cn}] '{text}'")
            except:
                pass

        if found_texts:
            print(f"✓ {len(found_texts)}개 텍스트 발견")
        else:
            print("✗ 텍스트 없음")

        return {
            "success": False,
            "message": "9가지 방법 모두 실패"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
