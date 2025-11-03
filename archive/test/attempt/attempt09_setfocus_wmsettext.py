"""
Attempt 09: 포커스 설정 + WM_SETTEXT + 변경 알림

SetFocus → WM_SETTEXT → EN_CHANGE 알림 시뮬레이션
"""
import time
import win32api
import win32con
import win32gui


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 09: 포커스 + WM_SETTEXT + 변경 알림")
    print("=" * 60)

    try:
        # 초기 상태 캡처
        capture_func("attempt09_00_initial.png")

        # 기본사항 탭으로 이동
        print("\n[1/5] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt09_01_basic_tab.png")

        # Edit 컨트롤 찾기
        print("\n[2/5] Edit 컨트롤 찾기...")
        edit_controls = []
        for ctrl in dlg.descendants():
            try:
                if "SPR32DU80EditHScroll" in ctrl.class_name():
                    text = ctrl.window_text()
                    edit_controls.append((ctrl, text))
                    print(f"  발견: HWND=0x{ctrl.handle:08X} text=\"{text}\"")
            except:
                pass

        if not edit_controls:
            return {"success": False, "message": "SPR32DU80 Edit 컨트롤을 찾을 수 없음"}

        print(f"  총 {len(edit_controls)}개 발견")

        # 테스트 데이터
        test_data = [
            ("사번", "2025001"),
            ("주민번호", "900101-1234567"),
            ("성명", "홍길동"),
        ]

        print(f"\n[3/5] 테스트 데이터: {test_data}")

        # EN_CHANGE 알림 메시지
        EN_CHANGE = 0x0300
        WM_COMMAND = 0x0111

        print("\n[4/5] 데이터 입력 시도...")
        for idx, (ctrl, original_text) in enumerate(edit_controls):
            if idx >= len(test_data):
                break

            try:
                label, new_text = test_data[idx]
                hwnd = ctrl.handle

                print(f"\n  [{idx}] {label} 입력 시도: \"{new_text}\"")
                print(f"      HWND=0x{hwnd:08X}, 원래 값=\"{original_text}\"")

                # 1. 포커스 설정
                print("      단계 1: SetFocus()")
                try:
                    win32gui.SetFocus(hwnd)
                    time.sleep(0.1)
                    print("      ✓ 포커스 설정됨")
                except Exception as e:
                    print(f"      ⚠ 포커스 설정 오류: {e}")

                # 2. 기존 텍스트 전체 선택
                print("      단계 2: EM_SETSEL (전체 선택)")
                EM_SETSEL = 0x00B1
                win32api.SendMessage(hwnd, EM_SETSEL, 0, -1)
                time.sleep(0.05)

                # 3. WM_SETTEXT로 텍스트 설정
                print(f"      단계 3: WM_SETTEXT(\"{new_text}\")")
                result = win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, new_text)
                print(f"      반환값: {result}")
                time.sleep(0.1)

                # 4. EN_CHANGE 알림 전송 (부모에게)
                print("      단계 4: EN_CHANGE 알림")
                try:
                    parent_hwnd = win32gui.GetParent(hwnd)
                    if parent_hwnd:
                        # HIWORD = EN_CHANGE, LOWORD = control ID
                        ctrl_id = win32api.GetWindowLong(hwnd, win32con.GWL_ID)
                        wparam = (EN_CHANGE << 16) | ctrl_id
                        win32api.SendMessage(parent_hwnd, WM_COMMAND, wparam, hwnd)
                        print(f"      ✓ 부모(0x{parent_hwnd:08X})에게 알림 전송")
                except Exception as e:
                    print(f"      ⚠ 알림 전송 오류: {e}")

                # 5. Enter 키 전송
                print("      단계 5: Enter 키")
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                time.sleep(0.05)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                time.sleep(0.1)

                # 결과 확인
                result_text = ctrl.window_text()
                print(f"      결과: \"{result_text}\"")
                if result_text == new_text:
                    print("      ✅ 성공!")
                else:
                    print(f"      ⚠ 불일치 (예상: \"{new_text}\", 실제: \"{result_text}\")")

            except Exception as e:
                print(f"  [{idx}] ✗ 오류: {e}")
                import traceback
                traceback.print_exc()

        time.sleep(1)
        capture_func("attempt09_02_after_input.png")

        # 최종 결과 확인
        print("\n[5/5] 최종 결과 확인...")
        success_count = 0
        for idx, (ctrl, original_text) in enumerate(edit_controls):
            if idx >= len(test_data):
                break
            try:
                label, expected_text = test_data[idx]
                current_text = ctrl.window_text()
                match = "✅" if current_text == expected_text else "❌"
                print(f"  [{idx}] {label}: \"{original_text}\" → \"{current_text}\" {match}")
                if current_text == expected_text:
                    success_count += 1
            except:
                pass

        message = f"입력 완료: {success_count}/{len(test_data)}개 성공"
        print(f"\n{message}")

        return {
            "success": success_count > 0,
            "message": message
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}


if __name__ == "__main__":
    # 단독 실행 테스트
    from pywinauto import application
    import sys
    import os

    # 프로젝트 루트를 sys.path에 추가
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")
    hwnd = dlg.handle

    def capture_func(filename):
        capture_window(hwnd, filename)

    result = run(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
