"""
분납적용 자동화 디버그 버전 6

사원 행을 Enter 키로 확실히 선택
"""
import sys
import win32process
import win32gui
import win32con
import win32api
import time
from pywinauto import Application
from pywinauto.keyboard import send_keys
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


def close_info_dialogs():
    """정보 다이얼로그가 있으면 닫기"""
    time.sleep(0.3)
    send_keys("{ENTER}")  # 확인 버튼 클릭
    time.sleep(0.3)


def process_installment_debug(data_list):
    """디버그 모드 - 사원 선택 후 Enter"""
    print("\n" + "="*70)
    print("분납적용 자동화 디버그 모드 v6 (사원 선택 + Enter)")
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
        capture_screen("debug6_01_initial.png")
        time.sleep(0.5)

        # 4. 첫 번째 사원 선택
        print("\n[4단계] 사원 선택")

        # 왼쪽 스프레드에 포커스
        left_spread.set_focus()
        time.sleep(0.3)

        # Home 키로 맨 위로
        send_keys("{HOME}")
        time.sleep(0.3)

        # Down 키로 첫 번째 사원 (헤더 다음)
        send_keys("{DOWN}")
        time.sleep(0.5)
        capture_screen("debug6_02_after_down.png")

        # Enter 키로 사원 선택 확정 (이게 핵심!)
        print("  Enter로 사원 선택 확정")
        send_keys("{ENTER}")
        time.sleep(1.0)
        capture_screen("debug6_03_after_enter.png")
        print("✓ 사원 선택 완료")

        # 5. 소득세 입력 (다이얼로그가 열렸을 수 있음)
        print("\n[5단계] 소득세 입력")
        print(f"  입력할 값: {test_data[0]['소득세']}")

        send_keys(str(test_data[0]['소득세']))
        time.sleep(1.0)
        capture_screen("debug6_04_income_tax.png")
        print("  ✓ 소득세 입력 완료")

        # 6. Tab으로 다음 필드
        print("\n[6단계] Tab으로 지방소득세 필드 이동")
        send_keys("{TAB}")
        time.sleep(0.5)
        capture_screen("debug6_05_tab_to_local.png")

        # 7. 지방소득세 입력
        print("\n[7단계] 지방소득세 입력")
        print(f"  입력할 값: {test_data[0]['지방소득세']}")

        send_keys(str(test_data[0]['지방소득세']))
        time.sleep(1.0)
        capture_screen("debug6_06_local_tax.png")
        print("  ✓ 지방소득세 입력 완료")

        # 8. Enter로 확정
        print("\n[8단계] Enter로 저장")
        send_keys("{ENTER}")
        time.sleep(1.0)
        capture_screen("debug6_07_after_save.png")
        print("  ✓ Enter 완료")

        # 최종 상태
        time.sleep(1.0)
        capture_screen("debug6_08_final.png")

        print("\n" + "="*70)
        print("[완료] 스크린샷 8개 생성됨")
        print("  debug6_01_initial.png ~ debug6_08_final.png")
        print("="*70)

    except Exception as e:
        import traceback
        print(f"\n❌ 오류 발생: {e}")
        print(traceback.format_exc())


def main():
    """메인 함수"""
    print("="*70)
    print("분납적용 자동화 디버그 v6")
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
    print()

    # 디버그 실행
    process_installment_debug(data)

    print("\n완료! 스크린샷을 확인하세요.")


if __name__ == "__main__":
    main()
