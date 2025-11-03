"""
직원 정보 입력 + 메시지 모니터링

Attempt 09 성공 방법을 메시지 모니터링과 함께 실행
"""
import sys
import time
import threading
from datetime import datetime
from pywinauto import application
from advanced_message_monitor import AdvancedMessageMonitor
import win32api
import win32con
import win32gui

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')


def input_employee_data(monitor, dlg, test_data):
    """직원 정보 입력 (Attempt 09 방식)"""
    time.sleep(0.5)  # 모니터링 준비 대기

    print(f"\n{'':>17}▶ 직원 정보 입력 시작")
    print(f"{'':>17}   데이터: {test_data}")

    try:
        # 기본사항 탭으로 이동
        print(f"{'':>17}   [1/3] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)

        # Edit 컨트롤 찾기
        print(f"{'':>17}   [2/3] Edit 컨트롤 찾기...")
        edit_controls = []
        for ctrl in dlg.descendants():
            try:
                if "SPR32DU80EditHScroll" in ctrl.class_name():
                    text = ctrl.window_text()
                    edit_controls.append((ctrl, text))
            except:
                pass

        if not edit_controls:
            print(f"{'':>17}   ✗ Edit 컨트롤을 찾을 수 없음")
            return

        print(f"{'':>17}   ✓ {len(edit_controls)}개 발견")

        # EN_CHANGE 알림 메시지
        EN_CHANGE = 0x0300
        WM_COMMAND = 0x0111

        print(f"{'':>17}   [3/3] 데이터 입력...")
        for idx, (ctrl, original_text) in enumerate(edit_controls):
            if idx >= len(test_data):
                break

            try:
                label, new_text = test_data[idx]
                hwnd = ctrl.handle

                print(f"{'':>17}     • {label}: \"{new_text}\"")

                # WM_SETTEXT
                monitor.log_message(hwnd, win32con.WM_SETTEXT, 0, 0, "📤 SEND: ")
                win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, new_text)
                time.sleep(0.1)

                # EN_CHANGE 알림
                try:
                    parent_hwnd = win32gui.GetParent(hwnd)
                    if parent_hwnd:
                        ctrl_id = win32api.GetWindowLong(hwnd, win32con.GWL_ID)
                        wparam = (EN_CHANGE << 16) | ctrl_id
                        monitor.log_message(parent_hwnd, WM_COMMAND, wparam, hwnd, "📤 SEND: ")
                        win32api.SendMessage(parent_hwnd, WM_COMMAND, wparam, hwnd)
                except:
                    pass

                # Enter 키
                monitor.log_message(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0, "📤 SEND: ")
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                time.sleep(0.05)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                time.sleep(0.3)

            except Exception as e:
                print(f"{'':>17}     ✗ 오류: {e}")

        print(f"{'':>17}   ✅ 입력 완료")

    except Exception as e:
        print(f"{'':>17}   ✗ 오류: {e}")
        import traceback
        traceback.print_exc()


def main():
    print("=" * 80)
    print("직원 정보 입력 + 메시지 모니터링")
    print("=" * 80)

    # 1. 연결
    print("\n[1/3] 사원등록 윈도우 연결...")
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        main_hwnd = dlg.handle
        print(f"✓ 메인 윈도우: HWND=0x{main_hwnd:08X}")
    except Exception as e:
        print(f"✗ 연결 실패: {e}")
        return

    # 2. 모니터 설정
    log_file = f"test/employee_input_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    monitor = AdvancedMessageMonitor(target_hwnd=None, log_file=log_file)

    # WM_SETTEXT, WM_COMMAND 추가
    monitor.filter_messages.add(win32con.WM_SETTEXT)
    monitor.filter_messages.add(win32con.WM_COMMAND)
    monitor.filter_messages.add(win32con.WM_KEYDOWN)

    # 3. 테스트 데이터
    test_data = [
        ("사번", "2025001"),
        ("주민번호", "900101-1234567"),
        ("성명", "홍길동"),
    ]

    print(f"\n[2/3] 메시지 모니터링 + 직원 정보 입력")
    monitor.start()

    # 입력 스레드
    input_thread = threading.Thread(
        target=input_employee_data,
        args=(monitor, dlg, test_data)
    )

    input_thread.start()
    input_thread.join()

    time.sleep(1)
    monitor.stop()

    # 4. 결과 분석
    print(f"\n[3/3] 메시지 분석")
    print("=" * 80)

    messages = monitor.get_messages()

    # 메시지 타입별 분류
    msg_types = {}
    for msg in messages:
        msg_name = msg['msg_name']
        msg_types[msg_name] = msg_types.get(msg_name, 0) + 1

    print(f"\n총 {len(messages)}개 메시지:")
    for msg_name, count in sorted(msg_types.items()):
        print(f"  • {msg_name}: {count}개")

    print(f"\n💾 전체 로그: {log_file}")

    print("\n" + "=" * 80)
    print("✅ 완료!")
    print("=" * 80)

    print("\n💡 확인사항:")
    print("  1. test/image/attempt09_*.png - 스크린샷 확인")
    print(f"  2. {log_file} - 메시지 로그 확인")
    print("  3. 사원등록 프로그램에서 데이터 확인")


if __name__ == "__main__":
    main()
