"""
시도 78: 위치 기반으로 모든 컨트롤 클릭 및 복사

주민등록번호가 Static이나 다른 컨트롤일 가능성
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
    print("시도 78: 위치 기반 모든 컨트롤 복사")
    print("="*60)

    target_value = "XXXXXX-XXXXXXX"  # 마스킹된 예시값
    print(f"찾을 값: '{target_value}'")

    try:
        # 초기 상태 캡처
        capture_func("attempt78_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"왼쪽 스프레드 HWND: 0x{left_spread.handle:08X}")

        print("\n=== 1. 스프레드에서 사원 선택 ===")

        # 활성화
        dlg.set_focus()
        time.sleep(0.3)
        left_spread.set_focus()
        time.sleep(0.3)

        # 현재 셀 값 복사
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.1)
        time.sleep(0.3)

        empno = pyperclip.paste()
        print(f"✓ 사번: '{empno}'")

        capture_func("attempt78_01_after_spread_select.png")

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

        capture_func("attempt78_02_after_tab_click.png")

        print("\n=== 3. 모든 컨트롤 (Edit, Static 등) 수집 ===")

        # 관심 컨트롤 타입들
        interesting_classes = ['Edit', 'Static', 'MaskEdit', 'RichEdit', 'RichEdit20W', 'ComboBox']

        controls = []
        for desc in dlg.descendants():
            class_name = desc.class_name()

            # 관심있는 클래스 또는 'edit' 포함
            if class_name in interesting_classes or 'edit' in class_name.lower():
                try:
                    rect = desc.rectangle()
                    # 화면에 보이는 컨트롤만 (크기가 있는)
                    if rect.width() > 10 and rect.height() > 10:
                        controls.append({
                            'control': desc,
                            'class': class_name,
                            'rect': rect,
                            'hwnd': desc.handle
                        })
                except:
                    pass

        print(f"관심 컨트롤: {len(controls)}개")

        # 위치순 정렬
        controls.sort(key=lambda c: (c['rect'].top, c['rect'].left))

        for i, ctrl_info in enumerate(controls):
            rect = ctrl_info['rect']
            print(f"  {i}: {ctrl_info['class']} at ({rect.left}, {rect.top})")

        print("\n=== 4. 각 컨트롤에서 값 읽기 시도 ===")

        found_values = []

        for i, ctrl_info in enumerate(controls):
            ctrl = ctrl_info['control']
            class_name = ctrl_info['class']
            rect = ctrl_info['rect']

            print(f"\n{i}. {class_name}:")

            # 중앙점
            center_x = (rect.left + rect.right) // 2
            center_y = (rect.top + rect.bottom) // 2

            try:
                # 방법 1: window_text()
                try:
                    text1 = ctrl.window_text()
                    if text1:
                        print(f"  window_text: '{text1}'")
                        if target_value in text1:
                            found_values.append(('window_text', i, text1))
                            print(f"    ✓✓✓ 타겟 발견!")
                except:
                    pass

                # 방법 2: 클릭 후 복사 (Edit 계열만)
                if 'edit' in class_name.lower():
                    try:
                        # 클릭
                        win32api.SetCursorPos((center_x, center_y))
                        time.sleep(0.05)
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                        time.sleep(0.05)
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                        time.sleep(0.15)

                        # 전체 선택 및 복사
                        ctrl.type_keys("^a^c", pause=0.1)
                        time.sleep(0.2)

                        text2 = pyperclip.paste()
                        if text2 and text2 != "BEFORE":
                            print(f"  복사: '{text2}'")
                            if target_value in text2:
                                found_values.append(('copy', i, text2))
                                print(f"    ✓✓✓ 타겟 발견!")
                    except:
                        pass

            except Exception as e:
                print(f"  오류: {e}")

        capture_func("attempt78_03_after_all_checks.png")

        if found_values:
            print(f"\n✓✓✓ 타겟 값 발견! ({len(found_values)}개)")

            for method, index, value in found_values:
                ctrl_info = controls[index]
                print(f"\n발견:")
                print(f"  방법: {method}")
                print(f"  인덱스: {index}")
                print(f"  클래스: {ctrl_info['class']}")
                print(f"  값: '{value}'")

            return {
                "success": True,
                "message": f"""
🎉 '{target_value}' 발견!

사번: '{empno}'
발견 개수: {len(found_values)}개

하지만 이 방법은 포그라운드에서만 작동합니다.
"""
            }
        else:
            print(f"\n✗ '{target_value}'를 찾지 못함")

            # 긴 숫자 문자열이라도 출력
            print("\n=== 5. 발견된 긴 텍스트 ===")
            for i, ctrl_info in enumerate(controls):
                try:
                    text = ctrl_info['control'].window_text()
                    if text and len(text) > 8:
                        print(f"  {i}. ({ctrl_info['class']}) '{text}'")
                except:
                    pass

            return {
                "success": False,
                "message": f"검사 완료: {len(controls)}개 컨트롤, 타겟 값 없음"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
