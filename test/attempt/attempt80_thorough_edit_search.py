"""
시도 80: 철저한 Edit 컨트롤 검색

모든 Edit 관련 컨트롤을 찾아서 하나하나 클릭 및 복사
"""
import time
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api
import pyperclip


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 80: 철저한 Edit 컨트롤 검색")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt80_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"왼쪽 스프레드 HWND: 0x{left_spread.handle:08X}")

        print("\n=== 1. 스프레드에서 사원 클릭 (SendMessage) ===")

        spread_hwnd = left_spread.handle

        # 스프레드 클릭 (데이터 로드를 위해)
        click_x, click_y = 100, 50
        lparam = win32api.MAKELONG(click_x, click_y)

        win32api.SendMessage(spread_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.1)
        win32api.SendMessage(spread_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.5)  # 데이터 로드 대기

        print(f"✓ 스프레드 클릭 완료: ({click_x}, {click_y})")

        # 활성화 후 사번 확인
        dlg.set_focus()
        time.sleep(0.3)
        left_spread.set_focus()
        time.sleep(0.3)

        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.1)
        time.sleep(0.3)

        empno = pyperclip.paste()
        print(f"✓ 사번 확인: '{empno}'")

        print("\n=== 2. 기본사항 탭 클릭 ===")

        # 탭 컨트롤 찾기
        tab_control = None
        for ctrl in dlg.descendants():
            if ctrl.class_name().startswith("Afx:TabWnd:"):
                tab_control = ctrl
                break

        if tab_control:
            tab_hwnd = tab_control.handle
            tab_x, tab_y = 50, 15
            tab_lparam = win32api.MAKELONG(tab_x, tab_y)

            win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tab_lparam)
            time.sleep(0.1)
            win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONUP, 0, tab_lparam)
            time.sleep(0.5)

            print("✓ 기본사항 탭 클릭 완료")

        capture_func("attempt80_01_after_tab_click.png")

        print("\n=== 3. Win32 API로 모든 Edit 컨트롤 찾기 ===")

        # WM_GETTEXT로 값 읽기
        def read_text_wm(hwnd):
            try:
                length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
                if length == 0:
                    return ""
                buffer = create_unicode_buffer(length + 1)
                win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
                return buffer.value
            except:
                return ""

        # 모든 자식 윈도우에서 Edit 찾기
        def find_all_edits(parent_hwnd):
            edits = []

            def enum_callback(hwnd, param):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    # Edit, MaskEdit, RichEdit 등
                    if 'Edit' in class_name or 'edit' in class_name.lower():
                        rect = win32gui.GetWindowRect(hwnd)
                        text_wm = read_text_wm(hwnd)
                        text_get = win32gui.GetWindowText(hwnd)

                        edits.append({
                            'hwnd': hwnd,
                            'class': class_name,
                            'rect': rect,
                            'text_wm': text_wm,
                            'text_get': text_get
                        })
                except:
                    pass
                return True

            win32gui.EnumChildWindows(parent_hwnd, enum_callback, None)
            return edits

        all_edits = find_all_edits(dlg.handle)
        print(f"Win32 API로 발견: {len(all_edits)}개")

        for i, edit in enumerate(all_edits):
            print(f"\n  Edit {i}:")
            print(f"    HWND: 0x{edit['hwnd']:08X}")
            print(f"    Class: {edit['class']}")
            print(f"    Rect: {edit['rect']}")
            print(f"    WM_GETTEXT: '{edit['text_wm']}'")
            print(f"    GetWindowText: '{edit['text_get']}'")

        print("\n=== 4. pywinauto로 모든 descendants 검색 ===")

        pywinauto_edits = []
        for desc in dlg.descendants():
            class_name = desc.class_name()
            if 'Edit' in class_name or 'edit' in class_name.lower():
                try:
                    pywinauto_edits.append({
                        'control': desc,
                        'class': class_name,
                        'hwnd': desc.handle,
                        'rect': desc.rectangle()
                    })
                except:
                    pass

        print(f"pywinauto로 발견: {len(pywinauto_edits)}개")

        for i, edit in enumerate(pywinauto_edits):
            rect = edit['rect']
            print(f"\n  Edit {i}:")
            print(f"    Class: {edit['class']}")
            print(f"    HWND: 0x{edit['hwnd']:08X}")
            print(f"    Rect: ({rect.left}, {rect.top}) - ({rect.right}, {rect.bottom})")

        print("\n=== 5. 각 Edit 컨트롤 클릭 및 복사 ===")

        found_values = []

        for i, edit_info in enumerate(pywinauto_edits):
            ctrl = edit_info['control']
            rect = edit_info['rect']

            print(f"\n  [{i}] {edit_info['class']}:")

            # 중앙점
            center_x = (rect.left + rect.right) // 2
            center_y = (rect.top + rect.bottom) // 2

            try:
                # 클릭
                win32api.SetCursorPos((center_x, center_y))
                time.sleep(0.1)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                time.sleep(0.2)

                # 전체 선택 및 복사
                pyperclip.copy("__EMPTY__")
                ctrl.type_keys("^a", pause=0.1)
                time.sleep(0.1)
                ctrl.type_keys("^c", pause=0.1)
                time.sleep(0.3)

                value = pyperclip.paste()

                if value and value != "__EMPTY__":
                    print(f"      ✓ 값: '{value}' (길이: {len(value)})")

                    # 숫자만 추출
                    digits = ''.join(c for c in value if c.isdigit())
                    print(f"        숫자: '{digits}' (길이: {len(digits)})")

                    if len(digits) == 13:
                        print(f"        ✓✓✓ 주민등록번호 형식!")
                        found_values.append({
                            'index': i,
                            'value': value,
                            'digits': digits
                        })
                    elif len(digits) >= 6:
                        print(f"        → 긴 숫자 발견")
                        found_values.append({
                            'index': i,
                            'value': value,
                            'digits': digits
                        })
                else:
                    print(f"      (비어있음)")

            except Exception as e:
                print(f"      오류: {e}")

        capture_func("attempt80_02_after_all_clicks.png")

        if found_values:
            print(f"\n✓✓✓ 값 발견: {len(found_values)}개")

            for item in found_values:
                print(f"\n  Edit {item['index']}:")
                print(f"    값: '{item['value']}'")
                print(f"    숫자: '{item['digits']}' (길이: {len(item['digits'])})")

                # 타겟 값과 비교 (예시)
                if len(item['digits']) == 13:
                    print(f"      ✓✓✓✓ 13자리 주민등록번호 형식!")

            return {
                "success": True,
                "message": f"""
발견: {len(found_values)}개 Edit 컨트롤에서 값 읽음

값들:
{chr(10).join([f"- '{v['value']}'" for v in found_values])}
"""
            }
        else:
            print("\n✗ 값을 읽을 수 있는 Edit 없음")

            return {
                "success": False,
                "message": f"Edit 검색 완료: Win32 {len(all_edits)}개, pywinauto {len(pywinauto_edits)}개, 값 읽기 실패"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
