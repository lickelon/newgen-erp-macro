"""
시도 105: 다양한 대안 방법들 시도

OCR 외의 모든 가능한 방법으로 스프레드 데이터 읽기
"""
import win32process
import win32gui
import win32con
import win32api
import time
import pyperclip
from pywinauto import Application
from ctypes import create_unicode_buffer


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 105: 다양한 대안 방법들 시도")
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
        right_spread = spreads[1] if len(spreads) > 1 else None

        left_hwnd = left_spread.handle
        print(f"\n왼쪽 스프레드: 0x{left_hwnd:08X}")
        if right_spread:
            right_hwnd = right_spread.handle
            print(f"오른쪽 스프레드: 0x{right_hwnd:08X}")

        # === 방법 1: 오른쪽 스프레드 시도 ===
        print("\n" + "="*60)
        print("[방법 1] 오른쪽 스프레드에서 복사 시도")
        print("="*60)

        if right_spread:
            right_spread.set_focus()
            time.sleep(0.3)

            pyperclip.copy("")
            win32api.SendMessage(right_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYUP, ord('C'), 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.3)

            value = pyperclip.paste()
            if value:
                print(f"✓✓✓ 성공! 오른쪽 스프레드 값: '{value}'")
                return {"success": True, "message": f"오른쪽 스프레드에서 읽기 성공: '{value}'"}
            else:
                print("✗ 오른쪽 스프레드도 복사 안 됨")

        # === 방법 2: 더블클릭으로 편집 모드 진입 ===
        print("\n" + "="*60)
        print("[방법 2] 더블클릭으로 편집 모드 진입 시도")
        print("="*60)

        left_spread.set_focus()
        time.sleep(0.2)

        # 셀 더블클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lParam)
        time.sleep(0.3)

        # Edit 컨트롤이 생성되었는지 확인
        edit_controls = []
        for child in installment_dlg.descendants():
            try:
                if child.class_name() == "Edit":
                    edit_controls.append(child)
            except:
                pass

        if edit_controls:
            print(f"✓ Edit 컨트롤 {len(edit_controls)}개 발견")
            for idx, edit in enumerate(edit_controls):
                try:
                    length = win32gui.SendMessage(edit.handle, win32con.WM_GETTEXTLENGTH, 0, 0)
                    if length > 0:
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(edit.handle, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value
                        print(f"  Edit[{idx}]: '{text}'")
                        if text and len(text) > 0:
                            print(f"✓✓✓ Edit 컨트롤에서 읽기 성공!")
                            return {"success": True, "message": f"Edit 컨트롤에서 읽기 성공: '{text}'"}
                except:
                    pass
        else:
            print("✗ Edit 컨트롤 생성 안 됨")

        # === 방법 3: Tab 키로 셀 간 이동 ===
        print("\n" + "="*60)
        print("[방법 3] Tab 키로 셀 간 이동 후 복사")
        print("="*60)

        left_spread.set_focus()
        time.sleep(0.2)

        for i in range(3):
            # Tab 키
            win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
            time.sleep(0.2)

            # 복사 시도
            pyperclip.copy("")
            win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('C'), 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.3)

            value = pyperclip.paste()
            if value:
                print(f"  Tab {i+1}회 후 복사 성공: '{value}'")
                return {"success": True, "message": f"Tab {i+1}회 후 복사 성공: '{value}'"}
            else:
                print(f"  Tab {i+1}회: 복사 안 됨")

        # === 방법 4: 스프레드의 모든 자식 컨트롤 확인 ===
        print("\n" + "="*60)
        print("[방법 4] 스프레드 자식 컨트롤 확인")
        print("="*60)

        try:
            children = left_spread.children()
            if children:
                print(f"✓ 자식 컨트롤 {len(children)}개 발견")
                for idx, child in enumerate(children):
                    try:
                        cn = child.class_name()
                        text = child.window_text()
                        print(f"  [{cn}] '{text}'")
                    except:
                        pass
            else:
                print("✗ 자식 컨트롤 없음")
        except:
            print("✗ 자식 컨트롤 조회 실패")

        # === 방법 5: 다양한 Windows 메시지 시도 ===
        print("\n" + "="*60)
        print("[방법 5] 다양한 Windows 메시지 시도")
        print("="*60)

        # LB_GETTEXT 시도 (리스트박스 메시지)
        LB_GETCOUNT = 0x018B
        LB_GETTEXT = 0x0189

        count = win32api.SendMessage(left_hwnd, LB_GETCOUNT, 0, 0)
        if count > 0:
            print(f"✓ LB_GETCOUNT: {count}개 항목")
            # 첫 번째 항목 읽기 시도
            buffer = create_unicode_buffer(256)
            result = win32api.SendMessage(left_hwnd, LB_GETTEXT, 0, buffer)
            if result > 0:
                print(f"✓✓✓ LB_GETTEXT 성공: '{buffer.value}'")
                return {"success": True, "message": f"LB_GETTEXT로 읽기 성공: '{buffer.value}'"}
        else:
            print("✗ LB_GETCOUNT: 0 (리스트박스 아님)")

        # === 방법 6: 스프레드 속성 직접 확인 ===
        print("\n" + "="*60)
        print("[방법 6] 스프레드 윈도우 속성 확인")
        print("="*60)

        # GetWindowLong으로 스타일 확인
        GWL_STYLE = -16
        GWL_EXSTYLE = -20

        style = win32api.GetWindowLong(left_hwnd, GWL_STYLE)
        ex_style = win32api.GetWindowLong(left_hwnd, GWL_EXSTYLE)

        print(f"Style: 0x{style:08X}")
        print(f"ExStyle: 0x{ex_style:08X}")

        # 읽기 전용인지 확인
        ES_READONLY = 0x0800
        if style & ES_READONLY:
            print("  → 읽기 전용(ES_READONLY) 플래그 있음")
        else:
            print("  → 읽기 전용 플래그 없음 (편집 가능?)")

        # === 방법 7: F2 키로 편집 모드 ===
        print("\n" + "="*60)
        print("[방법 7] F2 키로 편집 모드 진입")
        print("="*60)

        left_spread.set_focus()
        time.sleep(0.2)

        # F2 키 (엑셀에서 셀 편집 모드)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_F2, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_F2, 0)
        time.sleep(0.3)

        # Edit 컨트롤 다시 확인
        edit_controls = []
        for child in installment_dlg.descendants():
            try:
                if child.class_name() == "Edit":
                    edit_controls.append(child)
            except:
                pass

        if edit_controls:
            print(f"✓ F2 후 Edit 컨트롤 {len(edit_controls)}개")
            for idx, edit in enumerate(edit_controls):
                try:
                    length = win32gui.SendMessage(edit.handle, win32con.WM_GETTEXTLENGTH, 0, 0)
                    if length > 0:
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(edit.handle, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value
                        print(f"  Edit[{idx}]: '{text}'")
                        if text:
                            return {"success": True, "message": f"F2 후 Edit에서 읽기 성공: '{text}'"}
                except:
                    pass
        else:
            print("✗ F2 후에도 Edit 컨트롤 없음")

        # === 방법 8: 행 클릭 후 다른 컨트롤 확인 ===
        print("\n" + "="*60)
        print("[방법 8] 행 클릭 후 다른 컨트롤 변화 확인")
        print("="*60)

        # 행 클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.1)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
        time.sleep(0.5)

        print("✓ 행 클릭 완료")

        # 모든 Static/Edit 컨트롤 확인
        all_controls = []
        for ctrl in installment_dlg.descendants():
            try:
                cn = ctrl.class_name()
                if cn in ["Edit", "Static"]:
                    text = ctrl.window_text()
                    if text:
                        all_controls.append((cn, text))
            except:
                pass

        if all_controls:
            print(f"✓ {len(all_controls)}개 컨트롤에서 텍스트 발견:")
            for cn, text in all_controls[:10]:  # 처음 10개만
                print(f"  [{cn}] '{text}'")
        else:
            print("✗ 텍스트가 있는 컨트롤 없음")

        # === 방법 9: 다른 행 시도 ===
        print("\n" + "="*60)
        print("[방법 9] 여러 행 클릭 후 복사 시도")
        print("="*60)

        left_spread.set_focus()
        time.sleep(0.2)

        test_rows = [(50, 30), (50, 50), (50, 70), (50, 90)]

        for x, y in test_rows:
            # 클릭
            lParam = win32api.MAKELONG(x, y)
            win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
            time.sleep(0.05)
            win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
            time.sleep(0.2)

            # 복사 시도
            pyperclip.copy("")
            win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, ord('C'), 0)
            win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.3)

            value = pyperclip.paste()
            if value:
                print(f"✓✓✓ ({x}, {y}) 클릭 후 복사 성공: '{value}'")
                return {"success": True, "message": f"({x}, {y}) 위치 클릭 후 복사 성공: '{value}'"}
            else:
                print(f"✗ ({x}, {y}): 복사 안 됨")

        return {
            "success": False,
            "message": "9가지 방법 모두 실패 - fpUSpread80 데이터 읽기 불가"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
