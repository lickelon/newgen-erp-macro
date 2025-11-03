"""
시도 71: 기본사항 탭의 주민등록번호 읽기

아이디어:
- 스프레드에서 사원 선택 시 오른쪽 폼에 정보 로드됨
- 기본사항 탭의 주민등록번호 필드는 일반 Edit 컨트롤일 가능성 높음
- WM_GETTEXT로 백그라운드에서 읽을 수 있음!
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 71: 기본사항 탭의 주민등록번호 읽기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt71_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"왼쪽 스프레드 HWND: 0x{left_spread.handle:08X}")

        # 참조를 위해 현재 사원의 사번 복사로 확인
        import pyperclip
        left_spread.set_focus()
        time.sleep(0.3)
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.2)
        reference_empno = pyperclip.paste()
        print(f"참조 사번: '{reference_empno}'")

        print("\n=== 1. 기본사항 탭 찾기 ===")

        # 탭 컨트롤 찾기 (보통 SysTabControl32)
        def find_tabs(hwnd, tabs_list):
            class_name = win32gui.GetClassName(hwnd)
            if "tab" in class_name.lower():
                tabs_list.append(hwnd)
            win32gui.EnumChildWindows(hwnd, lambda h, l: find_tabs(h, l) or True, tabs_list)
            return True

        tabs = []
        win32gui.EnumChildWindows(dlg.handle, lambda h, l: find_tabs(h, l) or True, tabs)

        print(f"탭 컨트롤 {len(tabs)}개 발견:")
        for tab_hwnd in tabs:
            class_name = win32gui.GetClassName(tab_hwnd)
            text = win32gui.GetWindowText(tab_hwnd)
            print(f"  0x{tab_hwnd:08X}: {class_name} - '{text}'")

        print("\n=== 2. 모든 Edit 컨트롤 찾기 ===")

        def find_edits(hwnd, edits_list):
            class_name = win32gui.GetClassName(hwnd)
            if class_name == "Edit":
                edits_list.append(hwnd)
            win32gui.EnumChildWindows(hwnd, lambda h, l: find_edits(h, l) or True, edits_list)
            return True

        edits = []
        win32gui.EnumChildWindows(dlg.handle, lambda h, l: find_edits(h, l) or True, edits)

        print(f"Edit 컨트롤 {len(edits)}개 발견")

        # WM_GETTEXT로 값 읽기
        def read_edit_text(hwnd):
            length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            if length == 0:
                return ""

            buffer = create_unicode_buffer(length + 1)
            win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
            return buffer.value

        print("\n=== 3. Edit 컨트롤 값 읽기 ===")

        edit_values = {}
        for edit_hwnd in edits:
            try:
                value = read_edit_text(edit_hwnd)
                if value:  # 비어있지 않은 것만
                    rect = win32gui.GetWindowRect(edit_hwnd)
                    edit_values[edit_hwnd] = {
                        'value': value,
                        'rect': rect,
                        'length': len(value)
                    }
                    print(f"  0x{edit_hwnd:08X}: '{value}' (길이: {len(value)})")
            except Exception as e:
                pass

        # 주민등록번호 형식 찾기 (13자리 숫자 또는 6-7 형식)
        print("\n=== 4. 주민등록번호 형식 찾기 ===")

        resident_candidates = []
        for hwnd, info in edit_values.items():
            value = info['value']
            # 주민등록번호 패턴: 13자리 숫자 또는 XXXXXX-XXXXXXX
            if len(value) in [13, 14]:  # 13자리 또는 13+하이픈
                # 숫자만 추출
                digits_only = ''.join(c for c in value if c.isdigit())
                if len(digits_only) == 13:
                    resident_candidates.append({
                        'hwnd': hwnd,
                        'value': value,
                        'digits': digits_only
                    })
                    print(f"  ✓ 후보 발견: '{value}' (0x{hwnd:08X})")

        if not resident_candidates:
            print("  ✗ 주민등록번호 형식의 값을 찾을 수 없음")
            print("\n  모든 Edit 값:")
            for hwnd, info in sorted(edit_values.items(), key=lambda x: x[1]['length'], reverse=True):
                print(f"    '{info['value']}'")
        else:
            print(f"\n✓ 주민등록번호 후보 {len(resident_candidates)}개 발견")

            # 첫 번째 후보 사용
            resident_hwnd = resident_candidates[0]['hwnd']
            resident_value = resident_candidates[0]['value']

            print(f"\n선택된 주민등록번호: '{resident_value}'")

            # 백그라운드 테스트
            print("\n=== 5. 백그라운드 테스트 ===")
            print("메모장 실행...")
            notepad = subprocess.Popen(['notepad.exe'])
            time.sleep(2)

            active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            print(f"현재 활성 창: '{active}'")

            # 백그라운드에서 다시 읽기
            bg_value = read_edit_text(resident_hwnd)
            print(f"백그라운드 읽기: '{bg_value}'")

            notepad.terminate()

            if bg_value == resident_value:
                print("  ✓✓✓ 성공! 백그라운드에서도 같은 값 읽음")

                capture_func("attempt71_01_success.png")

                return {
                    "success": True,
                    "message": f"""
🎉 주민등록번호 백그라운드 읽기 성공!

주민등록번호: '{resident_value}'
참조 사번: '{reference_empno}'

이제 다음 로직 가능:
1. 스프레드에서 사원 선택 (클릭)
2. 주민등록번호 Edit 필드에서 WM_GETTEXT로 읽기 (백그라운드!)
3. 주민등록번호로 CSV 데이터 매칭
4. 부양가족 입력

이 방법은 완전히 백그라운드에서 작동합니다!
"""
                }
            else:
                print(f"  ✗ 값이 다름: 예상 '{resident_value}', 실제 '{bg_value}'")

        print("\n=== 6. 대안: 라벨 근처의 Edit 찾기 ===")

        # 모든 Static(라벨) 컨트롤 찾기
        def find_statics(hwnd, statics_list):
            class_name = win32gui.GetClassName(hwnd)
            if class_name == "Static":
                text = win32gui.GetWindowText(hwnd)
                if text:
                    rect = win32gui.GetWindowRect(hwnd)
                    statics_list.append({
                        'hwnd': hwnd,
                        'text': text,
                        'rect': rect
                    })
            win32gui.EnumChildWindows(hwnd, lambda h, l: find_statics(h, l) or True, statics_list)
            return True

        statics = []
        win32gui.EnumChildWindows(dlg.handle, lambda h, l: find_statics(h, l) or True, statics)

        # "주민등록번호" 라벨 찾기
        resident_labels = [s for s in statics if "주민" in s['text'] or "등록" in s['text']]

        if resident_labels:
            print(f"'주민등록번호' 관련 라벨 {len(resident_labels)}개 발견:")
            for label in resident_labels:
                print(f"  '{label['text']}' at {label['rect']}")

                # 라벨 오른쪽의 Edit 찾기
                label_rect = label['rect']
                for edit_hwnd, edit_info in edit_values.items():
                    edit_rect = edit_info['rect']
                    # 같은 줄에 있고 오른쪽에 있는 Edit
                    if abs(edit_rect[1] - label_rect[1]) < 30 and edit_rect[0] > label_rect[2]:
                        print(f"    → Edit 발견: '{edit_info['value']}'")

        capture_func("attempt71_01_complete.png")

        return {
            "success": False if not resident_candidates else True,
            "message": f"주민등록번호 후보 {len(resident_candidates)}개 발견, Edit 컨트롤 {len(edits)}개"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
