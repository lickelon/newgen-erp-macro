"""
시도 76: 특정 주민등록번호 값 찾기

pywinauto의 모든 방법을 동원하여 해당 값 검색
"""
import time
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 76: 특정 주민등록번호 값 찾기")
    print("="*60)

    target_value = "XXXXXX-XXXXXXX"  # 마스킹된 예시값
    print(f"찾을 값: '{target_value}'")

    try:
        # 초기 상태 캡처
        capture_func("attempt76_00_initial.png")

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

        capture_func("attempt76_01_after_clicks.png")

        print("\n=== 3. pywinauto로 모든 descendants 검색 ===")

        all_descendants = dlg.descendants()
        print(f"descendants: {len(all_descendants)}개")

        found_controls = []

        for i, desc in enumerate(all_descendants):
            try:
                # 여러 방법으로 텍스트 읽기
                texts_to_check = []

                # 1. window_text()
                try:
                    wt = desc.window_text()
                    if wt:
                        texts_to_check.append(('window_text', wt))
                except:
                    pass

                # 2. texts()
                try:
                    txts = desc.texts()
                    for txt in txts:
                        if txt:
                            texts_to_check.append(('texts', txt))
                except:
                    pass

                # 3. legacy_properties()
                try:
                    props = desc.legacy_properties()
                    if 'Value' in props and props['Value']:
                        texts_to_check.append(('Value', props['Value']))
                    if 'Name' in props and props['Name']:
                        texts_to_check.append(('Name', props['Name']))
                except:
                    pass

                # 4. element_info
                try:
                    elem_info = desc.element_info
                    if hasattr(elem_info, 'name') and elem_info.name:
                        texts_to_check.append(('element_info.name', elem_info.name))
                    if hasattr(elem_info, 'rich_text') and elem_info.rich_text:
                        texts_to_check.append(('element_info.rich_text', elem_info.rich_text))
                except:
                    pass

                # 타겟 값 포함 여부 확인
                for method, text in texts_to_check:
                    if target_value in text or text in target_value:
                        found_controls.append({
                            'index': i,
                            'control': desc,
                            'method': method,
                            'text': text,
                            'class': desc.class_name()
                        })
                        print(f"  ✓ 발견! [{i}] {method}: '{text}' ({desc.class_name()})")

            except Exception as e:
                pass

            # 진행상황 출력 (100개마다)
            if (i + 1) % 100 == 0:
                print(f"  ... {i + 1}/{len(all_descendants)} 검색 중")

        if found_controls:
            print(f"\n✓✓✓ '{target_value}' 포함 컨트롤 {len(found_controls)}개 발견!")

            for item in found_controls:
                ctrl = item['control']
                print(f"\n발견 #{item['index']}:")
                print(f"  방법: {item['method']}")
                print(f"  텍스트: '{item['text']}'")
                print(f"  클래스: {item['class']}")
                print(f"  HWND: 0x{ctrl.handle:08X}")

                # 이 컨트롤의 속성 상세 출력
                try:
                    print(f"  Rectangle: {ctrl.rectangle()}")
                except:
                    pass

                # 백그라운드 읽기 테스트
                print(f"\n  === 백그라운드 읽기 테스트 ===")

                # 메모장 실행
                import subprocess
                notepad = subprocess.Popen(['notepad.exe'])
                time.sleep(2)

                try:
                    # 다시 읽기
                    bg_text = ctrl.window_text()
                    print(f"  백그라운드 window_text(): '{bg_text}'")

                    if target_value in bg_text:
                        print(f"    ✓✓✓ 성공! 백그라운드에서도 읽기 가능!")

                        notepad.terminate()
                        capture_func("attempt76_02_success.png")

                        return {
                            "success": True,
                            "message": f"""
🎉 '{target_value}' 발견 및 백그라운드 읽기 성공!

발견 위치:
- 방법: {item['method']}
- 클래스: {item['class']}
- 텍스트: '{item['text']}'

백그라운드 읽기 가능! 🎊
"""
                        }
                except Exception as e:
                    print(f"  백그라운드 읽기 실패: {e}")

                notepad.terminate()

        else:
            print(f"\n✗ '{target_value}'를 찾을 수 없음")

            # 부분 일치 검색 (마지막 7자리만)
            print(f"\n=== 4. 부분 일치 검색 (마지막 7자리) ===")

            partial = target_value[-7:]  # 마스킹된 값의 마지막 7자리
            print(f"검색: '{partial}'")

            for i, desc in enumerate(all_descendants):
                try:
                    wt = desc.window_text()
                    if wt and partial in wt:
                        print(f"  부분 일치: '{wt}' ({desc.class_name()})")
                except:
                    pass

        capture_func("attempt76_02_complete.png")

        return {
            "success": len(found_controls) > 0,
            "message": f"검색 완료: descendants {len(all_descendants)}개, 발견 {len(found_controls)}개"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
