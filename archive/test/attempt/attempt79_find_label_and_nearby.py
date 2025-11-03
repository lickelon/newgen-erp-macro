"""
시도 79: "주민등록번호" 라벨 찾기 및 근처 컨트롤 읽기

라벨을 찾아서 그 오른쪽이나 아래에 있는 값 컨트롤 읽기
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
    print("시도 79: '주민등록번호' 라벨 찾기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt79_00_initial.png")

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

        capture_func("attempt79_01_after_tab_click.png")

        print("\n=== 3. '주민등록번호' 텍스트 검색 ===")

        # 모든 컨트롤에서 텍스트 검색
        all_controls = []
        for desc in dlg.descendants():
            try:
                text = desc.window_text()
                if text:
                    rect = desc.rectangle()
                    all_controls.append({
                        'control': desc,
                        'text': text,
                        'class': desc.class_name(),
                        'rect': rect,
                        'hwnd': desc.handle
                    })
            except:
                pass

        print(f"텍스트가 있는 컨트롤: {len(all_controls)}개")

        # "주민" 포함 검색
        jumin_controls = []
        for ctrl_info in all_controls:
            if '주민' in ctrl_info['text']:
                jumin_controls.append(ctrl_info)
                print(f"  ✓ 발견: '{ctrl_info['text']}' ({ctrl_info['class']})")

        if not jumin_controls:
            print("  ✗ '주민'을 포함한 텍스트 없음")

            # 전체 텍스트 일부 출력
            print("\n  발견된 텍스트 샘플 (처음 20개):")
            for ctrl_info in all_controls[:20]:
                text_preview = ctrl_info['text'][:30]
                print(f"    '{text_preview}' ({ctrl_info['class']})")

            return {
                "success": False,
                "message": "주민등록번호 라벨을 찾을 수 없음"
            }

        print(f"\n✓ '주민' 포함 컨트롤 {len(jumin_controls)}개 발견")

        print("\n=== 4. 라벨 근처의 컨트롤 찾기 ===")

        for jumin_ctrl in jumin_controls:
            label_rect = jumin_ctrl['rect']
            print(f"\n라벨: '{jumin_ctrl['text']}'")
            print(f"  위치: ({label_rect.left}, {label_rect.top}) - ({label_rect.right}, {label_rect.bottom})")

            # 라벨 오른쪽에 있는 컨트롤 찾기
            print("\n  오른쪽 컨트롤 검색:")

            nearby_controls = []
            for ctrl_info in all_controls:
                ctrl_rect = ctrl_info['rect']

                # 같은 줄 (세로 위치가 비슷함) + 오른쪽에 있음
                vertical_diff = abs(ctrl_rect.top - label_rect.top)
                if vertical_diff < 30 and ctrl_rect.left > label_rect.right:
                    distance = ctrl_rect.left - label_rect.right
                    nearby_controls.append({
                        'info': ctrl_info,
                        'distance': distance
                    })

            # 거리순 정렬 (가까운 것부터)
            nearby_controls.sort(key=lambda x: x['distance'])

            if nearby_controls:
                print(f"  발견: {len(nearby_controls)}개")

                for i, nearby in enumerate(nearby_controls[:5]):  # 처음 5개만
                    ctrl_info = nearby['info']
                    print(f"\n  [{i}] 거리: {nearby['distance']}px")
                    print(f"      클래스: {ctrl_info['class']}")
                    print(f"      텍스트: '{ctrl_info['text']}'")
                    print(f"      위치: ({ctrl_info['rect'].left}, {ctrl_info['rect'].top})")

                    # 13자리 숫자인지 확인
                    text = ctrl_info['text']
                    digits = ''.join(c for c in text if c.isdigit())
                    if len(digits) == 13:
                        print(f"      ✓✓✓ 주민등록번호 형식!")

            else:
                print("  오른쪽에 컨트롤 없음")

                # 아래쪽 검색
                print("\n  아래쪽 컨트롤 검색:")

                below_controls = []
                for ctrl_info in all_controls:
                    ctrl_rect = ctrl_info['rect']

                    # 비슷한 가로 위치 + 아래에 있음
                    horizontal_diff = abs(ctrl_rect.left - label_rect.left)
                    if horizontal_diff < 50 and ctrl_rect.top > label_rect.bottom:
                        distance = ctrl_rect.top - label_rect.bottom
                        below_controls.append({
                            'info': ctrl_info,
                            'distance': distance
                        })

                # 거리순 정렬
                below_controls.sort(key=lambda x: x['distance'])

                if below_controls:
                    print(f"  발견: {len(below_controls)}개")

                    for i, below in enumerate(below_controls[:5]):
                        ctrl_info = below['info']
                        print(f"\n  [{i}] 거리: {below['distance']}px")
                        print(f"      클래스: {ctrl_info['class']}")
                        print(f"      텍스트: '{ctrl_info['text']}'")

        print("\n=== 5. Edit/MaskEdit 컨트롤에서 복사 시도 ===")

        # Edit 계열 컨트롤만 필터링
        edit_controls = []
        for ctrl_info in all_controls:
            class_name = ctrl_info['class']
            if 'Edit' in class_name or 'edit' in class_name.lower():
                edit_controls.append(ctrl_info)

        print(f"Edit 계열 컨트롤: {len(edit_controls)}개")

        for i, ctrl_info in enumerate(edit_controls):
            ctrl = ctrl_info['control']
            rect = ctrl_info['rect']

            print(f"\n  Edit {i}: {ctrl_info['class']}")
            print(f"    위치: ({rect.left}, {rect.top})")

            # 클릭 후 복사
            try:
                center_x = (rect.left + rect.right) // 2
                center_y = (rect.top + rect.bottom) // 2

                win32api.SetCursorPos((center_x, center_y))
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                time.sleep(0.15)

                pyperclip.copy("EMPTY")
                ctrl.type_keys("^a^c", pause=0.1)
                time.sleep(0.2)

                value = pyperclip.paste()
                if value and value != "EMPTY":
                    print(f"    값: '{value}'")

                    digits = ''.join(c for c in value if c.isdigit())
                    if len(digits) == 13:
                        print(f"      ✓✓✓ 주민등록번호 형식! '{value}'")

            except Exception as e:
                print(f"    오류: {e}")

        capture_func("attempt79_02_complete.png")

        return {
            "success": len(jumin_controls) > 0,
            "message": f"'주민' 포함 라벨: {len(jumin_controls)}개, Edit 컨트롤: {len(edit_controls)}개"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
