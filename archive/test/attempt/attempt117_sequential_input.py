"""
시도 117: 순서 기반 자동 입력

사원코드 읽기 없이 CSV 순서대로 값 입력
"""
import win32process
import win32gui
import win32con
import win32api
import time
from pywinauto import Application
import csv


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 117: 순서 기반 자동 입력")
    print("="*60)

    # 테스트용 데이터 (실제로는 CSV에서 읽어옴)
    test_data = [
        {"사원코드": "2025021905", "사원명": "김지혜", "금액": "10000"},
        {"사원코드": "2025021906", "사원명": "오소희", "금액": "20000"},
        {"사원코드": "2025021907", "사원명": "이상호", "금액": "15000"},
    ]

    try:
        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[1단계] 급여자료입력 연결 (PID: {process_id})")

        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "#32770":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        installment_dlg = None
        dlg_hwnd = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            dlg_hwnd = hwnd
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 2. 스프레드 찾기
        print("\n[2단계] 스프레드 찾기")

        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        left_hwnd = left_spread.handle
        print(f"✓ 왼쪽 스프레드: 0x{left_hwnd:08X}")

        # 3. 첫 번째 사원으로 이동
        print("\n[3단계] 첫 번째 사원 선택")

        # Home 키로 맨 위로
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.3)

        # Down 키로 첫 번째 데이터 행으로 (헤더 스킵)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.3)

        print("✓ 첫 번째 사원 선택 완료")

        # 4. 각 사원에 대해 입력
        print(f"\n[4단계] {len(test_data)}명 데이터 입력")

        success_count = 0
        fail_count = 0

        for idx, row in enumerate(test_data):
            print(f"\n[{idx + 1}/{len(test_data)}] {row['사원명']} ({row['사원코드']})")

            try:
                # Tab 키로 오른쪽 입력 필드로 이동
                # (실제 필드 개수에 맞춰 조정 필요)
                for _ in range(3):
                    win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                    win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                    time.sleep(0.1)

                # 금액 입력
                amount = row['금액']
                print(f"  입력: {amount}")

                installment_dlg.type_keys(amount, with_spaces=False, pause=0.05)
                time.sleep(0.3)

                # Enter로 확정
                installment_dlg.type_keys("{ENTER}", pause=0.05)
                time.sleep(0.3)

                # Down 키로 다음 사원
                win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
                win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
                time.sleep(0.3)

                success_count += 1
                print(f"  ✓ 성공")

            except Exception as e:
                fail_count += 1
                print(f"  ✗ 실패: {e}")

        # 5. 결과
        print(f"\n[5단계] 입력 완료")
        print(f"  성공: {success_count}명")
        print(f"  실패: {fail_count}명")

        if success_count > 0:
            return {
                "success": True,
                "message": f"순서 기반 입력 성공: {success_count}명 완료 (실패: {fail_count}명)"
            }
        else:
            return {
                "success": False,
                "message": f"모든 입력 실패 ({fail_count}명)"
            }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
