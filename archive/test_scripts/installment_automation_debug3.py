"""
분납적용 자동화 디버그 버전 3

Tab 키를 다이얼로그 윈도우에 직접 전송
"""
import sys
import win32process
import win32gui
import win32con
import win32api
import time
from pywinauto import Application
import pandas as pd
from PIL import ImageGrab

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')


def load_yearend_data(excel_path):
    """연말정산 엑셀 파일 읽기"""
    print(f"\n[데이터 로드] {excel_path}")
    df = pd.read_excel(excel_path, header=None)
    print(f"✓ 총 {len(df)}행 로드")

    data = []
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        사원코드 = row[0]
        사원명 = row[1]
        소득세 = row[2]
        지방소득세 = row[3]

        if pd.isna(사원코드):
            continue

        if pd.isna(소득세):
            소득세 = 0
        if pd.isna(지방소득세):
            지방소득세 = 0

        data.append({
            "사원코드": str(사원코드).strip(),
            "사원명": str(사원명).strip(),
            "소득세": int(소득세) if 소득세 != 0 else 0,
            "지방소득세": int(지방소득세) if 지방소득세 != 0 else 0,
        })

    print(f"✓ {len(data)}명 데이터 파싱 완료")
    return data


def find_installment_dialog(app, process_id):
    """분납적용 다이얼로그 찾기"""
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

    for hwnd, title in found_dialogs:
        if not title:
            dialog = app.window(handle=hwnd)
            for child in dialog.children():
                try:
                    text = child.window_text()
                    if "분납적용" in text:
                        return dialog, hwnd
                except:
                    pass

    return None, None


def capture_screen(filename):
    """전체 화면 캡처"""
    img = ImageGrab.grab(all_screens=True)
    img.save(filename)
    print(f"  📷 스크린샷 저장: {filename}")


def process_installment_debug(data_list):
    """디버그 모드 - Tab을 다이얼로그에 전송"""
    print("\n" + "="*70)
    print("분납적용 자동화 디버그 모드 v3 (Tab to dialog)")
    print("="*70)

    # 첫 번째 사원만 처리
    test_data = data_list[:1]
    print(f"\n테스트 대상: {test_data[0]['사원명']}({test_data[0]['사원코드']})")
    print(f"  소득세: {test_data[0]['소득세']}")
    print(f"  지방소득세: {test_data[0]['지방소득세']}")

    try:
        # 1. 급여자료입력 연결
        print("\n[1단계] 급여자료입력 연결")
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)
        print(f"✓ PID: {process_id}")

        # 2. 분납적용 다이얼로그 찾기
        print("\n[2단계] 분납적용 다이얼로그 찾기")
        installment_dlg, dlg_hwnd = find_installment_dialog(app, process_id)

        if not installment_dlg:
            print("❌ 분납적용 다이얼로그를 찾지 못했습니다!")
            return

        print(f"✓ 분납적용 다이얼로그: 0x{dlg_hwnd:08X}")

        # 3. 왼쪽 스프레드 찾기
        print("\n[3단계] 스프레드 찾기")
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

        # 초기 상태 캡처
        capture_screen("debug3_01_initial.png")
        time.sleep(0.5)

        # 4. 첫 번째 사원 선택
        print("\n[4단계] 시작 위치 이동")

        # Home 키로 맨 위로
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.5)

        # Down 키로 첫 번째 사원 (헤더 다음)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
        time.sleep(0.5)
        capture_screen("debug3_02_first_employee.png")
        print("✓ 첫 번째 사원 선택 완료")

        # 5. Tab을 다이얼로그에 전송 (왼쪽 스프레드가 아닌 다이얼로그에)
        print("\n[5단계] Tab 키를 다이얼로그에 전송")

        for i in range(3):
            print(f"  Tab {i+1}/3 (다이얼로그 hwnd로)")
            win32api.SendMessage(dlg_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
            win32api.SendMessage(dlg_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
            time.sleep(0.5)
            capture_screen(f"debug3_03_tab_{i+1}.png")

        time.sleep(0.5)
        capture_screen("debug3_04_after_all_tabs.png")
        print("✓ Tab 전송 완료")

        # 6. 소득세 입력
        print("\n[6단계] 소득세 입력")
        print(f"  입력할 값: {test_data[0]['소득세']}")

        installment_dlg.type_keys(str(test_data[0]['소득세']), with_spaces=False, pause=0.1)
        time.sleep(1.0)
        capture_screen("debug3_05_income_tax.png")
        print("  ✓ 소득세 입력 완료")

        # 7. Tab으로 다음 필드
        print("\n[7단계] Tab으로 지방소득세 필드 이동")
        installment_dlg.type_keys("{TAB}", pause=0.1)
        time.sleep(0.5)
        capture_screen("debug3_06_tab_to_local.png")

        # 8. 지방소득세 입력
        print("\n[8단계] 지방소득세 입력")
        print(f"  입력할 값: {test_data[0]['지방소득세']}")

        installment_dlg.type_keys(str(test_data[0]['지방소득세']), with_spaces=False, pause=0.1)
        time.sleep(1.0)
        capture_screen("debug3_07_local_tax.png")
        print("  ✓ 지방소득세 입력 완료")

        # 9. Enter로 확정
        print("\n[9단계] Enter로 확정")
        installment_dlg.type_keys("{ENTER}", pause=0.1)
        time.sleep(1.0)
        capture_screen("debug3_08_after_enter.png")
        print("  ✓ Enter 완료")

        # 최종 상태
        time.sleep(1.0)
        capture_screen("debug3_09_final.png")

        print("\n" + "="*70)
        print("[완료] 스크린샷 9개 생성됨")
        print("  debug3_01_initial.png ~ debug3_09_final.png")
        print("="*70)

    except Exception as e:
        import traceback
        print(f"\n❌ 오류 발생: {e}")
        print(traceback.format_exc())


def main():
    """메인 함수"""
    print("="*70)
    print("분납적용 자동화 디버그 v3")
    print("="*70)

    # 데이터 로드
    excel_path = "연말정산.xls"
    data = load_yearend_data(excel_path)

    if not data:
        print("\n❌ 데이터가 없습니다!")
        return

    print(f"\n첫 번째 사원으로 디버그 테스트를 실행합니다.")
    print(f"  사원명: {data[0]['사원명']}")
    print(f"  사원코드: {data[0]['사원코드']}")
    print(f"  소득세: {data[0]['소득세']}")
    print(f"  지방소득세: {data[0]['지방소득세']}")

    response = input("\n계속하시겠습니까? (y/n): ")
    if response.lower() != 'y':
        print("취소되었습니다.")
        return

    # 디버그 실행
    process_installment_debug(data)

    print("\n완료! 스크린샷을 확인하세요.")


if __name__ == "__main__":
    main()
