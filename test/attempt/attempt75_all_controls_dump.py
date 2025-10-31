"""
시도 75: 모든 컨트롤 완전 덤프

사원 클릭 후 기본사항 탭에서 모든 컨트롤의 텍스트를 읽어서
주민등록번호가 어디에 있는지 찾기
"""
import time
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 75: 모든 컨트롤 완전 덤프")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt75_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        spread_hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{spread_hwnd:08X}")

        print("\n=== 1. 스프레드 클릭 ===")

        click_x, click_y = 100, 50
        lparam = win32api.MAKELONG(click_x, click_y)

        win32api.SendMessage(spread_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.1)
        win32api.SendMessage(spread_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(0.5)

        print("✓ 스프레드 클릭 완료")

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
        else:
            print("✗ 탭 컨트롤 없음")

        capture_func("attempt75_01_after_clicks.png")

        print("\n=== 3. 모든 컨트롤 텍스트 읽기 ===")

        # WM_GETTEXT로 읽기
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

        # GetWindowText로 읽기 (대안)
        def read_text_get(hwnd):
            try:
                return win32gui.GetWindowText(hwnd)
            except:
                return ""

        # 모든 컨트롤 수집
        all_controls = []

        def collect_controls(hwnd, controls_list):
            try:
                class_name = win32gui.GetClassName(hwnd)
                text_wm = read_text_wm(hwnd)
                text_get = read_text_get(hwnd)
                rect = win32gui.GetWindowRect(hwnd)

                controls_list.append({
                    'hwnd': hwnd,
                    'class': class_name,
                    'text_wm': text_wm,
                    'text_get': text_get,
                    'rect': rect
                })

                win32gui.EnumChildWindows(hwnd, lambda h, l: collect_controls(h, l) or True, controls_list)
            except:
                pass
            return True

        win32gui.EnumChildWindows(dlg.handle, lambda h, l: collect_controls(h, l) or True, all_controls)

        print(f"전체 컨트롤: {len(all_controls)}개")

        # 텍스트가 있는 컨트롤만 필터링
        controls_with_text = []
        for ctrl in all_controls:
            if ctrl['text_wm'] or ctrl['text_get']:
                text = ctrl['text_wm'] or ctrl['text_get']
                controls_with_text.append({
                    'hwnd': ctrl['hwnd'],
                    'class': ctrl['class'],
                    'text': text,
                    'method': 'WM_GETTEXT' if ctrl['text_wm'] else 'GetWindowText'
                })

        print(f"텍스트가 있는 컨트롤: {len(controls_with_text)}개\n")

        # 클래스별로 그룹화
        by_class = {}
        for ctrl in controls_with_text:
            cls = ctrl['class']
            if cls not in by_class:
                by_class[cls] = []
            by_class[cls].append(ctrl)

        # 클래스별 출력
        for cls, ctrls in sorted(by_class.items()):
            print(f"\n[{cls}] {len(ctrls)}개:")
            for ctrl in ctrls[:20]:  # 각 클래스당 최대 20개
                text_preview = ctrl['text'][:50] if len(ctrl['text']) > 50 else ctrl['text']
                print(f"  0x{ctrl['hwnd']:08X}: '{text_preview}' (길이: {len(ctrl['text'])})")

        print("\n=== 4. 주민등록번호 형식 검색 ===")

        # 13자리 숫자 찾기
        resident_candidates = []
        for ctrl in controls_with_text:
            text = ctrl['text']
            # 13자리 숫자 또는 하이픈 포함
            if len(text) in range(10, 20):  # 10-20자 사이
                digits_only = ''.join(c for c in text if c.isdigit())
                if len(digits_only) == 13:
                    resident_candidates.append(ctrl)
                    print(f"  ✓ 후보: '{text}' ({ctrl['class']})")

        if not resident_candidates:
            print("  주민등록번호 형식(13자리)을 찾을 수 없음")

            # "주민" 키워드로 검색
            print("\n=== 5. '주민' 키워드 검색 ===")

            jumin_related = [c for c in controls_with_text if '주민' in c['text']]
            if jumin_related:
                print(f"'주민' 포함 컨트롤: {len(jumin_related)}개")
                for ctrl in jumin_related:
                    print(f"  {ctrl['class']}: '{ctrl['text']}'")
            else:
                print("'주민' 키워드를 포함한 컨트롤 없음")

            # 긴 숫자 문자열 검색
            print("\n=== 6. 긴 숫자 문자열 검색 ===")

            numeric_texts = []
            for ctrl in controls_with_text:
                text = ctrl['text']
                digits = ''.join(c for c in text if c.isdigit())
                if len(digits) >= 6:  # 6자리 이상 숫자
                    numeric_texts.append({
                        'ctrl': ctrl,
                        'digits': digits,
                        'length': len(digits)
                    })

            if numeric_texts:
                print(f"6자리 이상 숫자 포함: {len(numeric_texts)}개")
                for item in sorted(numeric_texts, key=lambda x: x['length'], reverse=True)[:10]:
                    ctrl = item['ctrl']
                    print(f"  ({item['length']}자리) '{ctrl['text']}' ({ctrl['class']})")
            else:
                print("긴 숫자 문자열 없음")

        else:
            print(f"\n✓ 주민등록번호 후보 {len(resident_candidates)}개 발견")
            resident_ctrl = resident_candidates[0]
            print(f"선택: '{resident_ctrl['text']}' ({resident_ctrl['class']})")

        capture_func("attempt75_02_complete.png")

        return {
            "success": len(resident_candidates) > 0,
            "message": f"""
컨트롤 분석 완료:
- 전체 컨트롤: {len(all_controls)}개
- 텍스트 있음: {len(controls_with_text)}개
- 주민등록번호 후보: {len(resident_candidates)}개

클래스 종류: {len(by_class)}가지
"""
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
