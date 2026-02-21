"""
분납적용 자동화 모듈

연말정산 엑셀 파일에서 총액/1차분납/2차분납 데이터를 읽어서
분납적용 다이얼로그의 스프레드에 자동으로 입력합니다.
"""

import time
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List
import win32process
import win32gui
import xlrd
from pywinauto import Application
from pywinauto.keyboard import send_keys

from src.hotkey_manager import HotkeyManager


def load_yearend_data(excel_path):
    """연말정산 엑셀 파일 읽기 (xlrd, euc-kr 인코딩)"""
    print(f"\n[데이터 로드] {excel_path}")
    wb = xlrd.open_workbook(excel_path, encoding_override='euc-kr')
    ws = wb.sheet_by_index(0)
    print(f"✓ 총 {ws.nrows}행 로드")

    def safe_int(value):
        """셀 값을 int로 변환 (빈 값/문자열은 0)"""
        if isinstance(value, float):
            return int(value)
        if isinstance(value, str) and value.strip():
            try:
                return int(float(value))
            except ValueError:
                return 0
        return 0

    data = []
    for idx in range(2, ws.nrows):
        사원코드 = str(ws.cell_value(idx, 0)).strip()
        if not 사원코드:
            continue

        사원명 = str(ws.cell_value(idx, 1)).strip()
        총액_소득세 = safe_int(ws.cell_value(idx, 2))
        총액_지방소득세 = safe_int(ws.cell_value(idx, 3))
        분납1_소득세 = safe_int(ws.cell_value(idx, 5))
        분납1_지방소득세 = safe_int(ws.cell_value(idx, 6))
        분납2_소득세 = safe_int(ws.cell_value(idx, 8))
        분납2_지방소득세 = safe_int(ws.cell_value(idx, 9))

        data.append({
            "사원코드": 사원코드,
            "사원명": 사원명,
            "총액_소득세": 총액_소득세,
            "총액_지방소득세": 총액_지방소득세,
            "분납1_소득세": 분납1_소득세,
            "분납1_지방소득세": 분납1_지방소득세,
            "분납2_소득세": 분납2_소득세,
            "분납2_지방소득세": 분납2_지방소득세,
        })

    print(f"✓ {len(data)}명 데이터 파싱 완료")
    return data


class InstallmentAutomation:
    """분납적용 자동화"""

    def __init__(self, excel_path: str, verbose: bool = False, global_delay: float = 1.0):
        """
        초기화 및 연결

        Args:
            excel_path: Excel 파일 경로 (연말정산.xls)
            verbose: True면 DEBUG 로그 출력, False면 숨김 (기본값: False)
            global_delay: 전역 지연 시간 배율 (0.5~2.0, 기본값: 1.0)
        """
        self.verbose = verbose
        self.global_delay = max(0.5, min(2.0, global_delay))

        print(f"초기화 중...")

        # 로그 설정
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = log_dir / f"installment_{timestamp}.log"
        print(f"  [OK] 로그 파일: {self.log_file}")

        # 핫키 관리자 초기화
        self.hotkey_manager = HotkeyManager(log_callback=self.log)
        print(f"  [OK] 핫키 관리자 초기화 (Pause 키 리스너 등록)")

        # Excel 데이터 로드
        self.data = load_yearend_data(excel_path)
        print(f"  [OK] Excel 로드: {len(self.data)}명 사원")

        # pywinauto 연결 - 급여자료입력 윈도우
        try:
            self.app = Application(backend='win32')
            self.app.connect(title="급여자료입력")
            self.main_window = self.app.window(title="급여자료입력")
            _, self.process_id = win32process.GetWindowThreadProcessId(self.main_window.handle)
            print(f"  [OK] 급여자료입력 프로그램 연결 (PID: {self.process_id})")
        except Exception as e:
            error_msg = "❌ 급여자료입력 창을 찾을 수 없습니다!\n\n"
            error_msg += "다음을 확인하세요:\n"
            error_msg += "1. 급여자료입력 프로그램이 실행되어 있는지\n"
            error_msg += "2. 창 제목이 정확히 '급여자료입력'인지"
            raise Exception(error_msg)

        # 분납적용 다이얼로그 찾기
        self.dialog = None
        self.dialog_hwnd = None
        self.right_spread = None

    def log(self, level: str, message: str):
        """
        로그 기록

        Args:
            level: 로그 레벨 (INFO, SUCCESS, WARNING, ERROR, DEBUG)
            message: 메시지
        """
        # DEBUG 로그는 verbose 모드일 때만 출력
        if level == "DEBUG" and not self.verbose:
            return

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] {level}: {message}"

        # 콘솔 출력
        print(log_line)

        # 파일 기록
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(log_line + '\n')

    def check_stop_key(self) -> bool:
        """
        중지 요청 상태 체크

        Returns:
            True면 중지 요청, False면 계속 진행
        """
        return self.hotkey_manager.is_stop_requested()

    def stop(self):
        """중지 요청"""
        self.hotkey_manager.request_stop()

    def cleanup(self):
        """리소스 정리"""
        self.hotkey_manager.cleanup()

    def find_installment_dialog(self):
        """분납적용 다이얼로그 찾기"""
        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == self.process_id:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "#32770":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        for hwnd, title in found_dialogs:
            if not title:
                dialog = self.app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            self.dialog = dialog
                            self.dialog_hwnd = hwnd
                            return True
                    except:
                        pass

        return False

    def find_right_spread(self):
        """오른쪽 스프레드 찾기"""
        if not self.dialog:
            return False

        spreads = []
        for child in self.dialog.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        if len(spreads) < 2:
            return False

        # 왼쪽부터 정렬
        spreads.sort(key=lambda s: s.rectangle().left)
        self.right_spread = spreads[1]  # 오른쪽 스프레드 (입력 필드)
        return True

    def _type_and_enter(self, value: str):
        """값 입력 후 ENTER"""
        send_keys(value)
        time.sleep(0.3 * self.global_delay)
        send_keys("{ENTER}")
        time.sleep(0.3 * self.global_delay)

    def _skip_enter(self):
        """ENTER로 셀 스킵 (농특세 등)"""
        send_keys("{ENTER}")
        time.sleep(0.15 * self.global_delay)

    def process_one_employee(self, row: Dict) -> str:
        """
        한 명의 사원 데이터 입력

        Args:
            row: 사원 데이터 딕셔너리

        Returns:
            "installment" (분납 입력) 또는 "checkbox" (단건+체크) 또는 None (실패)
        """
        try:
            총액_소득세 = row['총액_소득세']

            if 총액_소득세 > 100000:
                # Case 1: 분납 대상 (총액 소득세 > 100,000)
                # 총액 소득세 → ENTER
                self._type_and_enter(str(row['총액_소득세']))
                # 총액 지방소득세 → ENTER
                self._type_and_enter(str(row['총액_지방소득세']))
                # 농특세 스킵 → ENTER
                self._skip_enter()
                # 1차분납 소득세 → ENTER
                self._type_and_enter(str(row['분납1_소득세']))
                # 1차분납 지방소득세 → ENTER
                self._type_and_enter(str(row['분납1_지방소득세']))
                # 1차분납 농특세 스킵 → ENTER
                self._skip_enter()
                # 2차분납 소득세 → ENTER
                self._type_and_enter(str(row['분납2_소득세']))
                # 2차분납 지방소득세 → ENTER
                self._type_and_enter(str(row['분납2_지방소득세']))
                # 2차분납 농특세 스킵 → ENTER → 3차분납 자동 → 다음 사원 자동이동
                self._skip_enter()
                return "installment"
            else:
                # Case 2: 분납 미대상 (총액 소득세 <= 100,000)
                # 총액 소득세 → ENTER
                self._type_and_enter(str(row['총액_소득세']))
                # 총액 지방소득세 → ENTER
                self._type_and_enter(str(row['총액_지방소득세']))
                # DOWN → 다음 사원 (농특세 셀에서)
                send_keys("{DOWN}")
                time.sleep(0.15 * self.global_delay)
                # LEFT → 지방소득세
                send_keys("{LEFT}")
                time.sleep(0.15 * self.global_delay)
                # LEFT → 소득세 셀로 이동
                send_keys("{LEFT}")
                time.sleep(0.15 * self.global_delay)
                return "skip"

        except Exception as e:
            self.log("ERROR", f"입력 실패: {e}")
            return None

    def run_checkbox(self, start_index: int = 0, count: int = None, dry_run: bool = False) -> Dict:
        """
        체크박스 체크 전용 실행 (<=100k 사원만 체크)

        체크박스 셀을 선택한 상태에서 시작해야 합니다.

        Args:
            start_index: 시작 인덱스 (기본값: 0)
            count: 처리할 사원 수 (None이면 전체)
            dry_run: True면 실제 입력 안함 (테스트용)

        Returns:
            결과 요약 딕셔너리
        """
        start_time = time.time()

        if count is None:
            count = len(self.data) - start_index

        process_data = self.data[start_index:start_index + count]

        self.log("INFO", "=== 체크박스 체크 시작 ===")
        self.log("INFO", f"처리 범위: {start_index + 1}번째 ~ {start_index + count}번째 사원")

        check_targets = [d for d in process_data if d['총액_소득세'] <= 100000]
        self.log("INFO", f"체크 대상: {len(check_targets)}명 (<=100k)")

        if dry_run:
            self.log("INFO", "!!! DRY RUN 모드 - 실제 입력 안함 !!!")

        # 분납적용 다이얼로그 찾기
        self.log("INFO", "분납적용 다이얼로그 찾기...")
        if not self.find_installment_dialog():
            self.log("ERROR", "❌ 분납적용 다이얼로그를 찾을 수 없습니다!")
            return {
                'status': 'error',
                'reason': 'dialog_not_found',
                'success': 0,
                'fail': 0,
                'skip': 0
            }
        self.log("INFO", f"✓ 다이얼로그 찾음: 0x{self.dialog_hwnd:08X}")

        # 스프레드 찾기
        self.log("INFO", "스프레드 컨트롤 찾기...")
        if not self.find_right_spread():
            self.log("ERROR", "❌ 스프레드를 찾을 수 없습니다!")
            return {
                'status': 'error',
                'reason': 'spread_not_found',
                'success': 0,
                'fail': 0,
                'skip': 0
            }
        self.log("INFO", f"✓ 오른쪽 스프레드 찾음: 0x{self.right_spread.handle:08X}")

        # 오른쪽 스프레드에 포커스
        self.log("INFO", "오른쪽 스프레드에 포커스 설정...")
        try:
            self.right_spread.set_focus()
            time.sleep(0.5 * self.global_delay)
            self.log("INFO", "✓ 포커스 설정 완료")
        except Exception as e:
            self.log("WARNING", f"⚠ 포커스 설정 실패: {e}")

        self.log("INFO", "⚠️ 중요: 첫 번째 사원의 체크박스 셀을 선택한 상태여야 합니다!")

        success_count = 0
        fail_count = 0
        skip_count = 0

        for idx, row in enumerate(process_data):
            if self.check_stop_key():
                self.log("WARNING", "중지 요청으로 처리 중단")
                break

            needs_check = row['총액_소득세'] <= 100000

            if needs_check:
                self.log("INFO", f"[{idx + 1}/{len(process_data)}] {row['사원명']} - 체크")
                if not dry_run:
                    try:
                        # SPACE → 체크박스 체크
                        send_keys("{SPACE}")
                        time.sleep(0.15 * self.global_delay)
                        # DOWN → 다음 사원
                        send_keys("{DOWN}")
                        time.sleep(0.15 * self.global_delay)
                        success_count += 1
                        self.log("SUCCESS", f"  [OK] 체크 완료")
                    except Exception as e:
                        fail_count += 1
                        self.log("ERROR", f"  [X] 체크 실패: {e}")
                else:
                    self.log("INFO", f"  → DRY RUN")
                    skip_count += 1
            else:
                self.log("DEBUG", f"[{idx + 1}/{len(process_data)}] {row['사원명']} - 스킵 (>100k)")
                if not dry_run:
                    # DOWN → 다음 사원 (체크 안하고 넘기기)
                    send_keys("{DOWN}")
                    time.sleep(0.15 * self.global_delay)

        elapsed = time.time() - start_time

        self.log("INFO", "=== 체크박스 체크 완료 ===")
        self.log("INFO", f"체크: {success_count}명")
        self.log("INFO", f"실패: {fail_count}명")
        self.log("INFO", f"소요 시간: {elapsed:.1f}초")

        return {
            'status': 'completed',
            'success': success_count,
            'fail': fail_count,
            'skip': skip_count,
            'elapsed': elapsed
        }

    def run(self, start_index: int = 0, count: int = None, dry_run: bool = False) -> Dict:
        """
        전체 자동화 실행

        Args:
            start_index: 시작 인덱스 (기본값: 0)
            count: 처리할 사원 수 (None이면 전체)
            dry_run: True면 실제 입력 안함 (테스트용)

        Returns:
            결과 요약 딕셔너리
        """
        start_time = time.time()

        if count is None:
            count = len(self.data) - start_index

        process_data = self.data[start_index:start_index + count]

        self.log("INFO", "=== 분납적용 자동화 시작 ===")
        self.log("INFO", f"처리 범위: {start_index + 1}번째 ~ {start_index + count}번째 사원")
        self.log("INFO", f"총 {count}명")

        if dry_run:
            self.log("INFO", "!!! DRY RUN 모드 - 실제 입력 안함 !!!")

        # 분납적용 다이얼로그 찾기
        self.log("INFO", "분납적용 다이얼로그 찾기...")
        if not self.find_installment_dialog():
            self.log("ERROR", "❌ 분납적용 다이얼로그를 찾을 수 없습니다!")
            return {
                'status': 'error',
                'reason': 'dialog_not_found',
                'success': 0,
                'fail': 0,
                'skip': 0
            }

        self.log("INFO", f"✓ 다이얼로그 찾음: 0x{self.dialog_hwnd:08X}")

        # 스프레드 찾기
        self.log("INFO", "스프레드 컨트롤 찾기...")
        if not self.find_right_spread():
            self.log("ERROR", "❌ 스프레드를 찾을 수 없습니다!")
            return {
                'status': 'error',
                'reason': 'spread_not_found',
                'success': 0,
                'fail': 0,
                'skip': 0
            }

        self.log("INFO", f"✓ 오른쪽 스프레드 찾음: 0x{self.right_spread.handle:08X}")

        # 오른쪽 스프레드에 포커스
        self.log("INFO", "오른쪽 스프레드에 포커스 설정...")
        try:
            self.right_spread.set_focus()
            time.sleep(0.5 * self.global_delay)
            self.log("INFO", "✓ 포커스 설정 완료")
        except Exception as e:
            self.log("WARNING", f"⚠ 포커스 설정 실패: {e}")

        self.log("INFO", "⚠️ 중요: 첫 번째 사원의 총액 소득세 셀을 선택한 상태여야 합니다!")
        self.log("INFO", "입력 시작!")

        # 각 사원 처리
        success_count = 0
        fail_count = 0
        skip_count = 0

        for idx, row in enumerate(process_data):
            # 중지 요청 체크
            if self.check_stop_key():
                self.log("WARNING", "중지 요청으로 처리 중단")
                break

            is_installment = row['총액_소득세'] > 100000
            self.log("INFO", f"[{idx + 1}/{len(process_data)}] {row['사원명']} ({row['사원코드']})")
            if is_installment:
                self.log("INFO", f"  분납 입력 | 총액: {row['총액_소득세']:,}/{row['총액_지방소득세']:,}, 1차: {row['분납1_소득세']:,}/{row['분납1_지방소득세']:,}, 2차: {row['분납2_소득세']:,}/{row['분납2_지방소득세']:,}")
            else:
                self.log("INFO", f"  단건 입력 | 총액: {row['총액_소득세']:,}/{row['총액_지방소득세']:,}")

            if not dry_run:
                result = self.process_one_employee(row)
                if result:
                    success_count += 1
                    self.log("SUCCESS", f"  [OK] {result}")
                else:
                    fail_count += 1
                    self.log("ERROR", f"  [X] 입력 실패")
            else:
                self.log("INFO", f"  → DRY RUN (실제 입력 안함)")
                skip_count += 1

        elapsed = time.time() - start_time

        self.log("INFO", "=== 입력 완료 ===")
        self.log("INFO", f"성공: {success_count}명")
        self.log("INFO", f"실패: {fail_count}명")
        if dry_run:
            self.log("INFO", f"건너뜀: {skip_count}명")
        self.log("INFO", f"소요 시간: {elapsed:.1f}초")

        return {
            'status': 'completed',
            'success': success_count,
            'fail': fail_count,
            'skip': skip_count,
            'elapsed': elapsed
        }
