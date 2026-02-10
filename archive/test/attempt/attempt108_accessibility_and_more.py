"""
시도 108: Accessibility API 및 추가 방법들

- IAccessible 인터페이스
- 자식 윈도우 상세 분석
- 숨겨진 컨트롤 찾기
- 메시지 훅 시도
"""
import win32process
import win32gui
import win32con
import win32api
import time
import pyperclip
from pywinauto import Application
from ctypes import create_unicode_buffer, windll, POINTER, c_long, byref
from comtypes import IUnknown, GUID
import comtypes.client


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 108: Accessibility API 및 추가 방법들")
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

        # === 방법 1: IAccessible 인터페이스 ===
        print("\n" + "="*60)
        print("[방법 1] IAccessible 인터페이스로 접근")
        print("="*60)

        try:
            from ctypes import oledll
            from comtypes import GUID

            # AccessibleObjectFromWindow
            OBJID_CLIENT = 0xFFFFFFFC
            IID_IAccessible = GUID("{618736E0-3C3D-11CF-810C-00AA00389B71}")

            # 함수 포인터 설정
            AccessibleObjectFromWindow = oledll.oleacc.AccessibleObjectFromWindow

            # IAccessible 가져오기
            pacc = POINTER(IUnknown)()
            result = AccessibleObjectFromWindow(
                left_hwnd,
                OBJID_CLIENT,
                byref(IID_IAccessible),
                byref(pacc)
            )

            if result == 0 and pacc:
                print("✓ IAccessible 인터페이스 획득 성공!")

                try:
                    # QueryInterface로 IAccessible로 변환
                    from comtypes.gen import Accessibility
                    accessible = pacc.QueryInterface(Accessibility.IAccessible)

                    # 이름 가져오기
                    try:
                        name = accessible.accName(0)
                        print(f"  이름: '{name}'")
                    except:
                        print("  이름: (없음)")

                    # 값 가져오기
                    try:
                        value = accessible.accValue(0)
                        print(f"  값: '{value}'")
                        if value:
                            print(f"\n✓✓✓ IAccessible.accValue 성공!")
                            return {"success": True, "message": f"IAccessible로 읽기 성공: '{value}'"}
                    except:
                        print("  값: (없음)")

                    # 자식 수 확인
                    try:
                        child_count = accessible.accChildCount
                        print(f"  자식 수: {child_count}")

                        if child_count > 0:
                            # 첫 번째 자식의 이름/값 확인
                            try:
                                child_name = accessible.accName(1)
                                child_value = accessible.accValue(1)
                                print(f"  자식[1] 이름: '{child_name}'")
                                print(f"  자식[1] 값: '{child_value}'")

                                if child_value:
                                    print(f"\n✓✓✓ 자식에서 값 발견!")
                                    return {"success": True, "message": f"자식 IAccessible: '{child_value}'"}
                            except:
                                pass
                    except:
                        print("  자식 수: (확인 불가)")

                except Exception as e:
                    print(f"  IAccessible 쿼리 실패: {e}")

            else:
                print(f"✗ AccessibleObjectFromWindow 실패: 0x{result:X}")

        except Exception as e:
            print(f"✗ IAccessible 시도 실패: {e}")

        # === 방법 2: 자식 윈도우 상세 분석 ===
        print("\n" + "="*60)
        print("[방법 2] 자식 윈도우 상세 분석 (모든 정보)")
        print("="*60)

        child_windows = []

        def enum_child_proc(hwnd, lParam):
            try:
                cn = win32gui.GetClassName(hwnd)

                # WM_GETTEXT 시도
                length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
                text = ""
                if length > 0:
                    buffer = create_unicode_buffer(length + 1)
                    win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
                    text = buffer.value

                # 스타일 확인
                style = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)

                # 사각형
                rect = win32gui.GetWindowRect(hwnd)

                child_windows.append({
                    'hwnd': hwnd,
                    'class': cn,
                    'text': text,
                    'style': style,
                    'rect': rect
                })
            except:
                pass
            return True

        win32gui.EnumChildWindows(left_hwnd, enum_child_proc, None)

        if child_windows:
            print(f"✓ 자식 윈도우 {len(child_windows)}개 상세 정보:")
            for idx, child in enumerate(child_windows):
                print(f"\n  [{idx}] HWND: 0x{child['hwnd']:08X}")
                print(f"      클래스: {child['class']}")
                print(f"      텍스트: '{child['text']}'")
                print(f"      스타일: 0x{child['style']:08X}")
                print(f"      위치: {child['rect']}")

                # 각 자식에 WM_GETTEXT 다시 시도 (더 큰 버퍼)
                try:
                    buffer = create_unicode_buffer(4096)
                    result = win32gui.SendMessage(child['hwnd'], win32con.WM_GETTEXT, 4096, buffer)
                    if result > 0:
                        text = buffer.value
                        print(f"      WM_GETTEXT(4KB): '{text[:100]}'")
                        if text and len(text) > len(child['text']):
                            print(f"\n✓✓✓ 큰 버퍼로 더 많은 데이터 발견!")
                            return {"success": True, "message": f"자식에서 읽기: '{text}'"}
                except:
                    pass
        else:
            print("✗ 자식 윈도우 없음")

        # === 방법 3: 숨겨진 리스트박스/리스트뷰 찾기 ===
        print("\n" + "="*60)
        print("[방법 3] 숨겨진 List 컨트롤 찾기")
        print("="*60)

        list_controls = []
        for ctrl in installment_dlg.descendants():
            try:
                cn = ctrl.class_name()
                if any(x in cn.lower() for x in ['list', 'tree', 'grid']):
                    list_controls.append((ctrl, cn))
                    print(f"  발견: [{cn}] HWND=0x{ctrl.handle:08X}")
            except:
                pass

        if list_controls:
            print(f"\n✓ List 계열 컨트롤 {len(list_controls)}개 발견")

            for ctrl, cn in list_controls:
                # LB_GETTEXT 시도 (리스트박스)
                if 'listbox' in cn.lower():
                    LB_GETCOUNT = 0x018B
                    LB_GETTEXT = 0x0189

                    count = win32api.SendMessage(ctrl.handle, LB_GETCOUNT, 0, 0)
                    print(f"  ListBox 항목 수: {count}")

                    if count > 0:
                        for i in range(min(count, 5)):  # 처음 5개만
                            buffer = create_unicode_buffer(256)
                            length = win32api.SendMessage(ctrl.handle, LB_GETTEXT, i, buffer)
                            if length > 0:
                                text = buffer.value
                                print(f"    [{i}]: '{text}'")
                                if text:
                                    return {"success": True, "message": f"ListBox에서 읽기: '{text}'"}
        else:
            print("✗ List 계열 컨트롤 없음")

        # === 방법 4: SendMessage 모든 일반 메시지 시도 ===
        print("\n" + "="*60)
        print("[방법 4] 다양한 표준 메시지 브루트포스")
        print("="*60)

        # 시도해볼 메시지들
        test_msgs = [
            (0x000D, "WM_GETTEXT"),
            (0x000E, "WM_GETTEXTLENGTH"),
            (0x0031, "WM_GETFONT"),
            (0x004E, "WM_NOTIFY"),
            (0x0111, "WM_COMMAND"),
            (0x018B, "LB_GETCOUNT"),
            (0x0189, "LB_GETTEXT"),
        ]

        found_responses = []
        for msg, desc in test_msgs:
            try:
                result = win32api.SendMessage(left_hwnd, msg, 0, 0)
                if result != 0:
                    found_responses.append((desc, result))
                    print(f"  {desc} (0x{msg:04X}): 0x{result:X}")
            except:
                pass

        if found_responses:
            print(f"✓ {len(found_responses)}개 메시지 응답 있음")
        else:
            print("✗ 모든 메시지 응답 없음")

        # === 방법 5: 스프레드 클릭 후 포커스된 컨트롤 확인 ===
        print("\n" + "="*60)
        print("[방법 5] 클릭 후 포커스된 컨트롤 추적")
        print("="*60)

        # 클릭 전 포커스
        before_focus = win32gui.GetFocus()
        print(f"클릭 전 포커스: 0x{before_focus:08X}")

        # 클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.05)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
        time.sleep(0.3)

        # 클릭 후 포커스
        after_focus = win32gui.GetFocus()
        print(f"클릭 후 포커스: 0x{after_focus:08X}")

        if after_focus != before_focus and after_focus != 0:
            print("✓ 포커스 변경됨!")

            # 새 포커스 컨트롤에서 텍스트 읽기
            try:
                cn = win32gui.GetClassName(after_focus)
                print(f"  새 포커스 클래스: {cn}")

                buffer = create_unicode_buffer(1024)
                length = win32gui.SendMessage(after_focus, win32con.WM_GETTEXT, 1024, buffer)
                if length > 0:
                    text = buffer.value
                    print(f"  텍스트: '{text}'")
                    if text:
                        return {"success": True, "message": f"포커스 변경 후 읽기: '{text}'"}
            except Exception as e:
                print(f"  읽기 실패: {e}")
        else:
            print("✗ 포커스 변경 없음")

        # === 방법 6: 다른 행들 순회하며 확인 ===
        print("\n" + "="*60)
        print("[방법 6] Down 키로 여러 행 순회하며 변화 감지")
        print("="*60)

        # Home으로 시작
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.2)

        # 3개 행 순회
        for row in range(3):
            print(f"\n  행 {row}:")

            # Down 키 (첫 행은 이미 선택되어 있으므로 스킵)
            if row > 0:
                win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
                win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
                time.sleep(0.3)

            # 모든 Edit/Static 확인
            for ctrl in installment_dlg.descendants():
                try:
                    cn = ctrl.class_name()
                    if cn in ["Edit", "Static"]:
                        length = win32gui.SendMessage(ctrl.handle, win32con.WM_GETTEXTLENGTH, 0, 0)
                        if length > 0 and length < 50:  # 짧은 텍스트만
                            buffer = create_unicode_buffer(length + 1)
                            win32gui.SendMessage(ctrl.handle, win32con.WM_GETTEXT, length + 1, buffer)
                            text = buffer.value

                            # 숫자나 코드 같은 패턴 확인
                            if text and (text.isdigit() or len(text) < 20):
                                print(f"    [{cn}] '{text}'")
                except:
                    pass

        return {
            "success": False,
            "message": "6가지 방법 모두 실패"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
