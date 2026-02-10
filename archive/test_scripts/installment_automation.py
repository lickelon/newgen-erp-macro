"""
분납적용 자동화 (최종 버전)

사용법:
1. 분납적용 다이얼로그를 엽니다
2. 오른쪽 스프레드의 첫 번째 사원, 소득세 열 셀을 클릭하여 선택합니다
3. 이 스크립트를 실행합니다
"""
import sys
import time
import win32process
import win32gui
from pywinauto import Application
from pywinauto.keyboard import send_keys
import pandas as pd

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


def find_installment_dialog():
    """분납적용 다이얼로그 찾기"""
    try:
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

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
    except:
        return None, None


def process_installment(data_list, start_index=0, count=None):
    """
    분납적용 데이터 입력

    Args:
        data_list: 입력할 데이터 리스트
        start_index: 시작 인덱스 (기본값: 0)
        count: 처리할 개수 (기본값: 전체)
    """
    if count is None:
        count = len(data_list) - start_index

    process_data = data_list[start_index:start_index + count]

    print("\n" + "="*70)
    print(f"분납적용 자동화 시작")
    print(f"  처리 범위: {start_index + 1}번째 ~ {start_index + count}번째 사원")
    print(f"  총 {count}명")
    print("="*70)

    # 다이얼로그 찾기
    print("\n[분납적용 다이얼로그 찾기]")
    dialog, dlg_hwnd = find_installment_dialog()

    if not dialog:
        print("❌ 분납적용 다이얼로그를 찾을 수 없습니다!")
        print("   다이얼로그가 열려있는지 확인하세요.")
        return 0, 0

    print(f"✓ 다이얼로그 찾음: 0x{dlg_hwnd:08X}")

    # 스프레드 찾기
    print("\n[스프레드 컨트롤 찾기]")
    spreads = []
    for child in dialog.children():
        try:
            if child.class_name() == "fpUSpread80":
                spreads.append(child)
        except:
            pass

    if len(spreads) < 2:
        print("❌ 스프레드를 찾을 수 없습니다!")
        return 0, 0

    # 왼쪽부터 정렬
    spreads.sort(key=lambda s: s.rectangle().left)
    right_spread = spreads[1]  # 오른쪽 스프레드 (입력 필드)

    print(f"✓ 오른쪽 스프레드 찾음: 0x{right_spread.handle:08X}")

    print("\n⚠️  중요: 오른쪽 스프레드의 첫 번째 사원, 소득세 열 셀을 선택한 상태여야 합니다!")
    print("5초 후 시작합니다...")
    for i in range(5, 0, -1):
        print(f"  {i}...")
        time.sleep(1)

    # 오른쪽 스프레드에 포커스
    print("\n[오른쪽 스프레드에 포커스 설정]")
    try:
        right_spread.set_focus()
        time.sleep(0.5)
        print("✓ 포커스 설정 완료")
    except Exception as e:
        print(f"⚠ 포커스 설정 실패: {e}")
        print("  수동으로 스프레드를 클릭해서 활성화해주세요.")
        time.sleep(2)

    print("\n입력 시작!\n")

    success_count = 0
    fail_count = 0

    for idx, row in enumerate(process_data):
        try:
            print(f"[{idx + 1}/{count}] {row['사원명']} ({row['사원코드']})")
            print(f"  소득세: {row['소득세']:,}, 지방소득세: {row['지방소득세']:,}")

            # 소득세 입력
            send_keys(str(row['소득세']))
            time.sleep(0.3)

            # Enter로 지방소득세 열로 이동
            send_keys("{ENTER}")
            time.sleep(0.3)

            # 지방소득세 입력
            send_keys(str(row['지방소득세']))
            time.sleep(0.3)

            # Enter로 확정
            send_keys("{ENTER}")
            time.sleep(0.3)

            # Down으로 다음 행으로 이동
            send_keys("{DOWN}")
            time.sleep(0.15)

            # Left 2번으로 소득세 열로 복귀
            send_keys("{LEFT}")
            time.sleep(0.15)
            send_keys("{LEFT}")
            time.sleep(0.15)

            success_count += 1

        except Exception as e:
            fail_count += 1
            print(f"  ❌ 오류: {e}")
            continue

    print("\n" + "="*70)
    print(f"입력 완료!")
    print(f"  성공: {success_count}명")
    print(f"  실패: {fail_count}명")
    print("="*70)

    return success_count, fail_count


def main():
    """메인 함수"""
    print("="*70)
    print("분납적용 자동화")
    print("="*70)

    # 데이터 로드
    excel_path = "연말정산.xls"
    data = load_yearend_data(excel_path)

    if not data:
        print("\n❌ 데이터가 없습니다!")
        return

    # 전체 실행 또는 테스트 실행 선택
    print(f"\n총 {len(data)}명의 데이터가 있습니다.")
    print("\n옵션을 선택하세요:")
    print("  1. 전체 실행 (3068명)")
    print("  2. 테스트 실행 (처음 10명만)")
    print("  3. 사용자 지정")

    choice = input("\n선택 (1/2/3): ").strip()

    if choice == "1":
        # 전체 실행
        process_installment(data)
    elif choice == "2":
        # 테스트 실행
        process_installment(data, start_index=0, count=10)
    elif choice == "3":
        # 사용자 지정
        start = int(input("시작 인덱스 (0부터 시작): "))
        count = int(input("처리할 개수: "))
        process_installment(data, start_index=start, count=count)
    else:
        print("잘못된 선택입니다.")


if __name__ == "__main__":
    main()
