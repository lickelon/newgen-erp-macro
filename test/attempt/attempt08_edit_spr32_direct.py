"""
Attempt 08: SPR32DU80 Edit 컨트롤에 직접 텍스트 입력

발견된 SPR32DU80EditHScroll 컨트롤에 직접 텍스트를 설정합니다.
"""
import time
import win32api
import win32con


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 08: SPR32DU80 Edit 컨트롤에 직접 텍스트 입력")
    print("=" * 60)

    try:
        # 초기 상태 캡처
        capture_func("attempt08_00_initial.png")

        # 기본사항 탭으로 이동
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt08_01_basic_tab.png")

        # Edit 컨트롤 찾기
        print("\n[2/4] Edit 컨트롤 찾기...")
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
        test_data = {
            "사번": "2025001",
            "성명": "홍길동",
            "주민번호": "900101-1234567",
            "나이": "35"
        }

        print(f"\n[3/4] 테스트 데이터: {test_data}")

        # 방법 1: set_edit_text() 사용
        print("\n[4/4] 데이터 입력 시도...")
        for idx, (ctrl, original_text) in enumerate(edit_controls):
            try:
                hwnd = ctrl.handle

                # 첫 번째 컨트롤에 사번 입력
                if idx == 0:
                    new_text = test_data["사번"]
                    print(f"\n  [{idx}] 사번 입력 시도: \"{new_text}\"")
                    print(f"      HWND=0x{hwnd:08X}, 원래 값=\"{original_text}\"")

                    # 방법 1: set_edit_text
                    print("      방법 1: set_edit_text()")
                    try:
                        ctrl.set_edit_text(new_text)
                        time.sleep(0.3)
                        result_text = ctrl.window_text()
                        print(f"      결과: \"{result_text}\"")
                        if result_text == new_text:
                            print("      ✓ 성공!")
                        else:
                            print("      ✗ 실패 (텍스트 불일치)")
                    except Exception as e:
                        print(f"      ✗ 오류: {e}")

                    # 방법 2: WM_SETTEXT
                    print("      방법 2: WM_SETTEXT")
                    try:
                        win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, new_text)
                        time.sleep(0.3)
                        result_text = ctrl.window_text()
                        print(f"      결과: \"{result_text}\"")
                        if result_text == new_text:
                            print("      ✓ 성공!")
                        else:
                            print("      ✗ 실패 (텍스트 불일치)")
                    except Exception as e:
                        print(f"      ✗ 오류: {e}")

                # 두 번째 컨트롤에 주민번호 입력
                elif idx == 1:
                    new_text = test_data["주민번호"]
                    print(f"\n  [{idx}] 주민번호 입력 시도: \"{new_text}\"")
                    print(f"      HWND=0x{hwnd:08X}, 원래 값=\"{original_text}\"")

                    try:
                        ctrl.set_edit_text(new_text)
                        time.sleep(0.3)
                        result_text = ctrl.window_text()
                        print(f"      결과: \"{result_text}\"")
                    except Exception as e:
                        print(f"      ✗ 오류: {e}")

                # 세 번째 컨트롤에 성명 입력
                elif idx == 2:
                    new_text = test_data["성명"]
                    print(f"\n  [{idx}] 성명 입력 시도: \"{new_text}\"")
                    print(f"      HWND=0x{hwnd:08X}, 원래 값=\"{original_text}\"")

                    try:
                        ctrl.set_edit_text(new_text)
                        time.sleep(0.3)
                        result_text = ctrl.window_text()
                        print(f"      결과: \"{result_text}\"")
                    except Exception as e:
                        print(f"      ✗ 오류: {e}")

            except Exception as e:
                print(f"  [{idx}] 오류: {e}")

        time.sleep(1)
        capture_func("attempt08_02_after_input.png")

        # 결과 확인
        print("\n결과 확인...")
        for idx, (ctrl, original_text) in enumerate(edit_controls):
            try:
                current_text = ctrl.window_text()
                print(f"  [{idx}] \"{original_text}\" → \"{current_text}\"")
            except:
                pass

        return {"success": True, "message": "SPR32DU80 Edit 입력 시도 완료 (결과는 스크린샷 확인)"}

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
    print(f"\n최종 결과: {result}")
