"""
시도 81: 하이픈 포함 주민등록번호 형태 검색

이전에는 숫자만 (13자리) 검색했으나
실제로는 "XXXXXX-XXXXXXX" (하이픈 포함 14자리) 형태일 수 있음
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
    print("시도 81: 하이픈 포함 주민등록번호 검색")
    print("="*60)

    target_with_hyphen = "XXXXXX-XXXXXXX"  # 마스킹된 예시값
    target_without_hyphen = "XXXXXXXXXXXXX"  # 마스킹된 예시값

    print(f"찾을 값:")
    print(f"  하이픈 포함: '{target_with_hyphen}'")
    print(f"  하이픈 없음: '{target_without_hyphen}'")

    try:
        # 초기 상태 캡처
        capture_func("attempt81_00_initial.png")

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

        # 사번 확인
        dlg.set_focus()
        time.sleep(0.3)
        left_spread.set_focus()
        time.sleep(0.3)

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

        capture_func("attempt81_01_after_clicks.png")

        print("\n=== 3. 모든 컨트롤에서 하이픈 포함 형태 검색 ===")

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

            # 하이픈 포함 형태 검색
            if target_with_hyphen in text:
                found.append({
                    'ctrl': ctrl_info,
                    'match_type': '하이픈 포함 완전 일치',
                    'text': text
                })
                print(f"  ✓✓✓ 발견 (하이픈 포함)! '{text}' ({ctrl_info['class']})")

            # 하이픈 없는 형태 검색
            elif target_without_hyphen in text:
                found.append({
                    'ctrl': ctrl_info,
                    'match_type': '하이픈 없음 완전 일치',
                    'text': text
                })
                print(f"  ✓✓✓ 발견 (하이픈 없음)! '{text}' ({ctrl_info['class']})")

            # 부분 일치는 실제 값으로 테스트할 때 사용
            # 예시 코드는 주석 처리
            # elif "XXXXXX" in text:  # 실제 검색할 앞 6자리로 교체
            #     found.append({
            #         'ctrl': ctrl_info,
            #         'match_type': '앞 6자리 부분 일치',
            #         'text': text
            #     })
            #     print(f"  ✓ 부분 일치: '{text}' ({ctrl_info['class']})")

        if found:
            print(f"\n✓✓✓ 발견! ({len(found)}개)")

            for item in found:
                ctrl_info = item['ctrl']
                print(f"\n매칭 타입: {item['match_type']}")
                print(f"  클래스: {ctrl_info['class']}")
                print(f"  HWND: 0x{ctrl_info['hwnd']:08X}")
                print(f"  텍스트: '{item['text']}'")

                # 백그라운드 테스트
                print(f"\n  === 백그라운드 읽기 테스트 ===")

                import subprocess
                notepad = subprocess.Popen(['notepad.exe'])
                time.sleep(2)

                active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                print(f"  현재 활성 창: '{active}'")

                # 백그라운드에서 다시 읽기
                try:
                    bg_text = ctrl_info['control'].window_text()
                    print(f"  백그라운드 읽기: '{bg_text}'")

                    if target_with_hyphen in bg_text or target_without_hyphen in bg_text:
                        print(f"    ✓✓✓ 성공! 백그라운드에서도 읽기 가능!")

                        notepad.terminate()
                        capture_func("attempt81_02_success.png")

                        return {
                            "success": True,
                            "message": f"""
🎉 주민등록번호 발견 및 백그라운드 읽기 성공!

매칭 타입: {item['match_type']}
클래스: {ctrl_info['class']}
텍스트: '{item['text']}'

백그라운드 읽기 가능! 🎊
"""
                        }
                except Exception as e:
                    print(f"  백그라운드 읽기 실패: {e}")

                notepad.terminate()

        else:
            print("\n✗ 발견 못함")

            # 하이픈이 포함된 다른 텍스트 찾기
            print("\n=== 4. 하이픈이 포함된 모든 텍스트 ===")

            hyphen_texts = [c for c in all_controls if '-' in c['text']]
            if hyphen_texts:
                print(f"하이픈 포함 텍스트: {len(hyphen_texts)}개")
                for ctrl_info in hyphen_texts[:20]:  # 처음 20개만
                    print(f"  '{ctrl_info['text']}' ({ctrl_info['class']})")
            else:
                print("하이픈 포함 텍스트 없음")

            # 6자리-7자리 패턴 찾기
            print("\n=== 5. XXXXXX-XXXXXXX 패턴 검색 ===")

            import re
            pattern = re.compile(r'\d{6}-\d{7}')

            pattern_matches = []
            for ctrl_info in all_controls:
                if pattern.search(ctrl_info['text']):
                    pattern_matches.append(ctrl_info)
                    print(f"  패턴 일치: '{ctrl_info['text']}' ({ctrl_info['class']})")

            if not pattern_matches:
                print("XXXXXX-XXXXXXX 패턴 없음")

        capture_func("attempt81_02_complete.png")

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
