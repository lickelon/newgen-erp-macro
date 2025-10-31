"""
시도 82: 설정 파일에서 민감정보 읽기 예제

하드코딩 대신 test_config.json에서 테스트 데이터를 읽어옵니다.
"""
import time
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 82: 설정 파일에서 민감정보 읽기")
    print("="*60)

    # ✅ 올바른 방법: 설정 파일에서 읽기
    try:
        from test.config import get_test_config

        config = get_test_config()

        # 설정에서 테스트 데이터 로드
        target_with_hyphen = config.resident_number_with_hyphen
        target_without_hyphen = config.resident_number_without_hyphen
        empno = config.empno

        print(f"✓ 설정 파일 로드 성공!")
        print(f"  사번: '{empno}'")
        print(f"  주민등록번호 (하이픈 포함): '{target_with_hyphen}'")
        print(f"  주민등록번호 (하이픈 없음): '{target_without_hyphen}'")

    except FileNotFoundError as e:
        return {
            "success": False,
            "message": f"""
❌ 설정 파일이 없습니다!

{e}

다음 명령으로 설정 파일을 만드세요:
  copy test_config.example.json test_config.json

그리고 test_config.json에 실제 테스트 데이터를 입력하세요.
"""
        }

    # ❌ 잘못된 방법 (하드코딩)
    # target_with_hyphen = "123456-1234567"  # 이렇게 하지 마세요!
    # empno = "0000000000"  # 이것도 안 됩니다!

    try:
        # 초기 상태 캡처
        capture_func("attempt82_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        spread_hwnd = left_spread.handle

        print(f"\n왼쪽 스프레드 HWND: 0x{spread_hwnd:08X}")

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

        capture_func("attempt82_01_after_clicks.png")

        print("\n=== 3. 설정 파일에서 읽어온 값으로 검색 ===")

        # 모든 descendants 검색
        all_controls = []
        for desc in dlg.descendants():
            try:
                text = desc.window_text()
                if text:
                    all_controls.append({
                        'control': desc,
                        'text': text,
                        'class': desc.class_name(),
                        'hwnd': desc.handle
                    })
            except:
                pass

        print(f"텍스트가 있는 컨트롤: {len(all_controls)}개")

        found = []

        for ctrl_info in all_controls:
            text = ctrl_info['text']

            # 설정 파일에서 읽어온 값으로 검색
            if target_with_hyphen in text:
                found.append({
                    'ctrl': ctrl_info,
                    'match_type': '하이픈 포함 완전 일치',
                    'text': text
                })
                print(f"  ✓✓✓ 발견 (하이픈 포함)! '{text}' ({ctrl_info['class']})")

            elif target_without_hyphen in text:
                found.append({
                    'ctrl': ctrl_info,
                    'match_type': '하이픈 없음 완전 일치',
                    'text': text
                })
                print(f"  ✓✓✓ 발견 (하이픈 없음)! '{text}' ({ctrl_info['class']})")

        if found:
            print(f"\n✓✓✓ 발견! ({len(found)}개)")

            for item in found:
                ctrl_info = item['ctrl']
                print(f"\n매칭 타입: {item['match_type']}")
                print(f"  클래스: {ctrl_info['class']}")
                print(f"  HWND: 0x{ctrl_info['hwnd']:08X}")
                print(f"  텍스트: '{item['text']}'")

            capture_func("attempt82_02_success.png")

            return {
                "success": True,
                "message": f"""
🎉 주민등록번호 발견!

매칭 타입: {found[0]['match_type']}
클래스: {found[0]['ctrl']['class']}
텍스트: '{found[0]['text']}'

✅ 민감정보를 하드코딩하지 않고 설정 파일에서 읽어왔습니다!
"""
            }
        else:
            print("\n✗ 발견 못함")

        capture_func("attempt82_02_complete.png")

        return {
            "success": len(found) > 0,
            "message": f"검색 완료: {len(all_controls)}개 컨트롤, 발견 {len(found)}개"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
