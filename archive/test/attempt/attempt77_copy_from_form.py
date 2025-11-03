"""
시도 77: 폼에서 주민등록번호 복사

1. 스프레드에서 사원 선택 (포그라운드 복사로 확인)
2. 기본사항 탭 클릭
3. 주민등록번호 필드 클릭 후 복사
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
    print("시도 77: 폼에서 주민등록번호 복사")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt77_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"왼쪽 스프레드 HWND: 0x{left_spread.handle:08X}")

        print("\n=== 1. 스프레드에서 사원 선택 (포그라운드) ===")

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
        print(f"✓ 선택된 사원 사번: '{empno}'")

        capture_func("attempt77_01_after_spread_select.png")

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
            capture_func("attempt77_02_after_tab_click.png")
        else:
            print("✗ 탭 컨트롤 없음")

        print("\n=== 3. 모든 Edit 컨트롤 위치 확인 ===")

        # Edit 컨트롤 찾기
        edits = []
        for desc in dlg.descendants():
            if desc.class_name() == 'Edit':
                rect = desc.rectangle()
                edits.append({
                    'control': desc,
                    'rect': rect,
                    'hwnd': desc.handle
                })

        print(f"Edit 컨트롤: {len(edits)}개")

        # 위치순 정렬 (위에서 아래, 왼쪽에서 오른쪽)
        edits.sort(key=lambda e: (e['rect'].top, e['rect'].left))

        for i, edit in enumerate(edits):
            rect = edit['rect']
            print(f"  Edit {i}: ({rect.left}, {rect.top}) - ({rect.right}, {rect.bottom})")

        if len(edits) == 0:
            return {"success": False, "message": "Edit 컨트롤을 찾을 수 없음"}

        print("\n=== 4. 각 Edit에 클릭 → 전체선택 → 복사 ===")

        results = []

        for i, edit in enumerate(edits):
            ctrl = edit['control']
            rect = edit['rect']
            hwnd = edit['hwnd']

            # 중앙점 계산
            center_x = (rect.left + rect.right) // 2
            center_y = (rect.top + rect.bottom) // 2

            print(f"\nEdit {i}:")
            print(f"  위치: ({center_x}, {center_y})")

            try:
                # 화면 절대 좌표로 클릭
                win32api.SetCursorPos((center_x, center_y))
                time.sleep(0.1)

                # 클릭
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                time.sleep(0.2)

                # 전체 선택
                ctrl.type_keys("^a", pause=0.1)
                time.sleep(0.1)

                # 복사
                pyperclip.copy("EMPTY")
                ctrl.type_keys("^c", pause=0.1)
                time.sleep(0.2)

                value = pyperclip.paste()

                if value and value != "EMPTY":
                    print(f"  ✓ 값: '{value}' (길이: {len(value)})")

                    # 13자리 숫자 체크
                    digits_only = ''.join(c for c in value if c.isdigit())
                    if len(digits_only) == 13:
                        print(f"    ✓✓ 주민등록번호 형식!")
                        results.append({
                            'index': i,
                            'value': value,
                            'digits': digits_only
                        })
                else:
                    print(f"  (비어있음)")

            except Exception as e:
                print(f"  ✗ 오류: {e}")

        capture_func("attempt77_03_after_copy_attempts.png")

        if results:
            print(f"\n✓✓✓ 주민등록번호 {len(results)}개 발견!")

            for result in results:
                print(f"  Edit {result['index']}: '{result['value']}'")

            # 첫 번째 결과 사용
            resident_number = results[0]['value']

            return {
                "success": True,
                "message": f"""
🎉 주민등록번호 발견!

사번: '{empno}'
주민등록번호: '{resident_number}'

하지만 이 방법은:
- ❌ 마우스 클릭 필요
- ❌ 포그라운드 필요
- ❌ 백그라운드 불가능

결론: 주민등록번호는 백그라운드에서 읽을 수 없음
→ 사번으로 매칭하는 것이 유일한 방법!
"""
            }
        else:
            print("\n✗ 주민등록번호를 찾지 못함")

            return {
                "success": False,
                "message": f"Edit {len(edits)}개 검사, 주민등록번호 형식 없음"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
