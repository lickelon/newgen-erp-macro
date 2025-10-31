"""
시도 73: 기본사항 탭 클릭 후 주민등록번호 읽기

docs/tab-automation.md의 방법 사용:
- 탭 컨트롤에 WM_LBUTTONDOWN/UP 전송
- 기본사항 탭 좌표: (50, 15)
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 73: 기본사항 탭 클릭 후 주민등록번호 읽기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt73_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        print(f"왼쪽 스프레드 HWND: 0x{left_spread.handle:08X}")

        # 참조 사번 확인 (활성화하지 않고 일단 스킵)
        reference_empno = "(확인 스킵)"
        print(f"참조 사번: {reference_empno}")

        print("\n=== 1. 탭 컨트롤 찾기 ===")

        # 탭 컨트롤 찾기 (Afx:TabWnd: 로 시작)
        tab_control = None
        for ctrl in dlg.descendants():
            if ctrl.class_name().startswith("Afx:TabWnd:"):
                tab_control = ctrl
                break

        if tab_control is None:
            return {"success": False, "message": "탭 컨트롤을 찾을 수 없음"}

        tab_hwnd = tab_control.handle
        print(f"✓ 탭 컨트롤 HWND: 0x{tab_hwnd:08X}")
        print(f"  클래스명: {tab_control.class_name()}")

        print("\n=== 2. 기본사항 탭 클릭 ===")

        # 기본사항 탭 좌표 (클라이언트 좌표)
        x, y = 50, 15

        # LPARAM 생성
        lparam = win32api.MAKELONG(x, y)

        # WM_LBUTTONDOWN 전송
        win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.1)

        # WM_LBUTTONUP 전송
        win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.5)  # 탭 전환 대기

        print("✓ 기본사항 탭 클릭 완료")
        capture_func("attempt73_01_after_tab_click.png")

        print("\n=== 3. 모든 컨트롤 다시 검색 ===")

        # WM_GETTEXT로 값 읽기
        def read_text(hwnd):
            try:
                length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
                if length == 0:
                    return ""
                buffer = create_unicode_buffer(length + 1)
                win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
                return buffer.value
            except:
                return ""

        # 모든 컨트롤 수집
        all_controls = []

        def collect_controls(hwnd, controls_list):
            try:
                class_name = win32gui.GetClassName(hwnd)
                text = read_text(hwnd)
                rect = win32gui.GetWindowRect(hwnd)

                controls_list.append({
                    'hwnd': hwnd,
                    'class': class_name,
                    'text': text,
                    'rect': rect
                })

                win32gui.EnumChildWindows(hwnd, lambda h, l: collect_controls(h, l) or True, controls_list)
            except:
                pass
            return True

        win32gui.EnumChildWindows(dlg.handle, lambda h, l: collect_controls(h, l) or True, all_controls)

        print(f"전체 컨트롤: {len(all_controls)}개")

        print("\n=== 4. Edit 컨트롤 값 읽기 ===")

        edits = [c for c in all_controls if c['class'] == 'Edit']
        print(f"Edit 컨트롤: {len(edits)}개")

        edit_values = {}
        for edit_ctrl in edits:
            if edit_ctrl['text']:
                edit_values[edit_ctrl['hwnd']] = edit_ctrl['text']
                print(f"  0x{edit_ctrl['hwnd']:08X}: '{edit_ctrl['text']}' (길이: {len(edit_ctrl['text'])})")

        print("\n=== 5. 주민등록번호 형식 찾기 ===")

        resident_candidates = []
        for hwnd, value in edit_values.items():
            # 주민등록번호 패턴: 13자리 숫자 또는 XXXXXX-XXXXXXX
            if len(value) in [13, 14]:
                digits_only = ''.join(c for c in value if c.isdigit())
                if len(digits_only) == 13:
                    resident_candidates.append({
                        'hwnd': hwnd,
                        'value': value,
                        'digits': digits_only
                    })
                    print(f"  ✓ 후보: '{value}' (0x{hwnd:08X})")

        if not resident_candidates:
            print("  ✗ 주민등록번호 형식을 찾을 수 없음")
            print("\n  긴 Edit 값들 (길이 내림차순):")
            sorted_edits = sorted(edit_values.items(), key=lambda x: len(x[1]), reverse=True)
            for hwnd, value in sorted_edits[:10]:
                print(f"    (길이 {len(value)}) '{value}'")

            # MaskEdit 등 다른 타입도 확인
            print("\n=== 6. 다른 컨트롤 타입 확인 ===")

            # MaskEdit, RichEdit 등
            mask_edits = [c for c in all_controls if 'edit' in c['class'].lower() or 'mask' in c['class'].lower()]
            print(f"\nMask/Edit 관련 컨트롤: {len(mask_edits)}개")
            for ctrl in mask_edits:
                if ctrl['text'] and ctrl['class'] != 'Edit':
                    print(f"  {ctrl['class']}: '{ctrl['text']}'")

        else:
            print(f"\n✓ 주민등록번호 후보 {len(resident_candidates)}개 발견")

            resident_hwnd = resident_candidates[0]['hwnd']
            resident_value = resident_candidates[0]['value']

            print(f"선택: '{resident_value}'")

            # 백그라운드 테스트
            print("\n=== 7. 백그라운드 테스트 ===")
            print("메모장 실행...")
            notepad = subprocess.Popen(['notepad.exe'])
            time.sleep(2)

            active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            print(f"현재 활성 창: '{active}'")

            # 백그라운드에서 다시 읽기
            bg_value = read_text(resident_hwnd)
            print(f"백그라운드 읽기: '{bg_value}'")

            notepad.terminate()

            if bg_value == resident_value:
                print("  ✓✓✓ 성공! 백그라운드에서도 같은 값!")

                capture_func("attempt73_02_success.png")

                return {
                    "success": True,
                    "message": f"""
🎉 주민등록번호 백그라운드 읽기 성공!

주민등록번호: '{resident_value}'
참조 사번: '{reference_empno}'

새로운 자동화 플로우:
1. 스프레드에서 사원 클릭
2. 기본사항 탭 선택 (WM_LBUTTONDOWN/UP)
3. 주민등록번호 읽기 (WM_GETTEXT) ← 백그라운드 가능!
4. 주민등록번호로 CSV 매칭
5. 부양가족 입력

완전한 백그라운드 자동화 가능! 🎊
"""
                }

        capture_func("attempt73_02_complete.png")

        return {
            "success": False if not resident_candidates else True,
            "message": f"Edit 컨트롤 {len(edits)}개, 주민등록번호 후보 {len(resident_candidates)}개"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
