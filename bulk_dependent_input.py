"""
부양가족 대량 입력 자동화 스크립트

CSV 파일에서 부양가족 데이터를 읽어서
왼쪽 사원 목록과 매칭하여 자동으로 입력합니다.
"""

import time
import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List
import pyperclip
import keyboard

from pywinauto import Application
from src.csv_reader import CSVReader, DependentData
from src.hotkey_manager import HotkeyManager
from src.spread_controller import SpreadController
from src.input_handler import InputHandler


class BulkDependentInput:
    """부양가족 대량 입력 자동화"""

    def __init__(self, csv_path: str, verbose: bool = False, global_delay: float = 1.0, start_from_current: bool = False):
        """
        초기화 및 연결

        Args:
            csv_path: CSV 파일 경로
            verbose: True면 DEBUG 로그 출력, False면 숨김 (기본값: False)
            global_delay: 전역 지연 시간 배율 (0.5~2.0, 기본값: 1.0)
            start_from_current: True면 현재 위치에서 시작, False면 Ctrl+Home 실행 (기본값: False)
        """
        self.verbose = verbose
        self.global_delay = max(0.5, min(2.0, global_delay))
        self.start_from_current = start_from_current

        print(f"초기화 중...")

        # 로그 설정 (먼저 설정)
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = log_dir / f"bulk_input_{timestamp}.log"
        print(f"  [OK] 로그 파일: {self.log_file}")

        # 핫키 관리자 초기화
        self.hotkey_manager = HotkeyManager(log_callback=self.log)
        print(f"  [OK] 핫키 관리자 초기화 (Pause 키 리스너 등록)")

        # CSV 데이터 로드
        self.csv_reader = CSVReader(csv_path)
        self.csv_reader.read()
        self.csv_data = self.csv_reader.group_by_employee()
        print(f"  [OK] CSV 로드: {len(self.csv_data)}명 사원")

        # pywinauto 연결
        try:
            self.app = Application(backend='win32')
            self.app.connect(title="사원등록")
            self.dlg = self.app.window(title="사원등록")
            print(f"  [OK] 사원등록 프로그램 연결")
        except Exception as e:
            error_msg = "❌ 사원등록 창을 찾을 수 없습니다!\n\n"
            error_msg += "다음을 확인하세요:\n"
            error_msg += "1. 사원등록 프로그램이 실행되어 있는지\n"
            error_msg += "2. 창 제목이 정확히 '사원등록'인지"
            raise Exception(error_msg)

        # 스프레드 컨트롤러 초기화 및 스프레드 찾기
        self.spread_controller = SpreadController(self.dlg, log_callback=self.log)
        try:
            self.left_spread, self.right_spread = self.spread_controller.find_spreads()
            print(f"  [OK] 스프레드 찾기 완료 (왼쪽/오른쪽)")

            # 왼쪽 포커스 설정
            self.left_spread.set_focus()
            time.sleep(0.3 * self.global_delay)
            print(f"  [OK] 왼쪽 사원 목록 포커스 설정")
        except Exception as e:
            raise

        # 입력 핸들러 초기화
        self.input_handler = InputHandler(
            self.left_spread,
            self.right_spread,
            global_delay=self.global_delay,
            log_callback=self.log
        )
        print(f"  [OK] 입력 핸들러 초기화")

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

    def read_current_employee_no(self) -> str:
        """
        현재 선택된 사원의 사번 읽기 (클립보드 방식)

        Returns:
            사번
        """
        start_time = time.time()

        # HOME → RIGHT → Ctrl+C → 사번
        self.input_handler.type_keys_with_delay(self.left_spread, "{HOME}", pause=0.05)
        self.input_handler.type_keys_with_delay(self.left_spread, "{RIGHT}", pause=0.05)
        employee_no = self.input_handler.copy_from_control()

        elapsed = time.time() - start_time
        self.log("DEBUG", f"⏱️ 사번 읽기: {elapsed:.2f}초")

        return employee_no

    def input_dependent(self, dep: DependentData) -> bool:
        """
        부양가족 한 명의 데이터 입력

        Args:
            dep: 부양가족 데이터

        Returns:
            성공 여부
        """
        start_time = time.time()

        try:
            # 1. 관계코드 입력
            t1 = time.time()
            self.input_handler.type_keys_with_delay(
                self.right_spread, dep.relationship_code, with_spaces=False, pause=0.05
            )
            self.log("DEBUG", f"  ⏱️ 관계코드: {time.time()-t1:.2f}초")

            # 2. 성명 입력 + ENTER
            t2 = time.time()
            self.input_handler.type_keys_with_delay(
                self.right_spread, dep.name, with_spaces=False, pause=0.05
            )
            self.input_handler.type_keys_with_delay(self.right_spread, "{ENTER}", pause=0.05)
            self.log("DEBUG", f"  ⏱️ 성명: {time.time()-t2:.2f}초")

            # 3. 내/외국인 입력
            t3 = time.time()
            nationality_code = self.input_handler.convert_nationality(dep.nationality)
            self.input_handler.type_keys_with_delay(
                self.right_spread, nationality_code, with_spaces=False, pause=0.05
            )
            self.log("DEBUG", f"  ⏱️ 내/외국인: {time.time()-t3:.2f}초")

            # 4. 번호 타입 입력
            t4 = time.time()
            id_type_code, id_number_clean = self.input_handler.clean_id_number(dep.id_number)
            self.input_handler.type_keys_with_delay(
                self.right_spread, id_type_code, with_spaces=False, pause=0.05
            )
            self.log("DEBUG", f"  ⏱️ 번호타입: {time.time()-t4:.2f}초")

            # 5. 주민등록번호/외국인등록번호/여권번호 입력
            t5 = time.time()
            self.input_handler.type_keys_with_delay(
                self.right_spread, id_number_clean, with_spaces=False, pause=0.05
            )
            self.log("DEBUG", f"  ⏱️ 번호입력: {time.time()-t5:.2f}초")

            # 6. 나이는 자동 입력되므로 건너뜀

            # 7. 기본공제여부
            t7 = time.time()
            is_basic_deduction = dep.basic_deduction.strip().upper() != 'N'
            if not is_basic_deduction:
                self.input_handler.type_keys_with_delay(
                    self.right_spread, "0", with_spaces=False, pause=0.05
                )
            else:
                self.input_handler.type_keys_with_delay(
                    self.right_spread, "{RIGHT}", with_spaces=False, pause=0.05
                )
            self.log("DEBUG", f"  ⏱️ 기본공제: {time.time()-t7:.2f}초")

            # 기본공제 Y인 경우에만 이후 필드 입력
            if is_basic_deduction:
                # 8. 장애유형 (0이면 건너뛰기)
                #    만나이 60 이상 + 기본공제 Y → 경로공제 열이 자동체크되므로 한 칸 더 이동
                t8 = time.time()
                try:
                    age = int(dep.age.strip())
                except ValueError:
                    age = 0
                if age >= 60:
                    self.input_handler.type_keys_with_delay(self.right_spread, "{RIGHT}", pause=0.05)

                if dep.disability_type.strip() == '0' or not dep.disability_type.strip():
                    self.input_handler.type_keys_with_delay(
                        self.right_spread, "{RIGHT}", pause=0.05
                    )
                else:
                    self.input_handler.type_keys_with_delay(
                        self.right_spread, dep.disability_type.strip(), with_spaces=False, pause=0.05
                    )
                self.log("DEBUG", f"  ⏱️ 장애유형: {time.time()-t8:.2f}초")

                # 9. 자녀공제 (연말관계가 4인 경우만)
                t9 = time.time()
                if dep.relationship_code == '4':
                    child_code = '1' if dep.child_deduction.strip().upper() == 'Y' else '0'
                    self.input_handler.type_keys_with_delay(
                        self.right_spread, child_code, with_spaces=False, pause=0.05
                    )
                    time.sleep(0.3)
                self.log("DEBUG", f"  ⏱️ 자녀공제: {time.time()-t9:.2f}초")

            # 다음 부양가족으로 이동
            t9 = time.time()
            self.input_handler.type_keys_with_delay(self.right_spread, "{HOME}", pause=0.05)
            self.input_handler.type_keys_with_delay(self.right_spread, "{DOWN}", pause=0.05)
            self.log("DEBUG", f"  ⏱️ 다음행이동: {time.time()-t9:.2f}초")

            elapsed = time.time() - start_time
            self.log("DEBUG", f"⏱️ 부양가족 1명 전체: {elapsed:.2f}초")

            return True

        except Exception as e:
            self.log("ERROR", f"input_dependent 실패: {e}")
            return False

    def clear_existing_dependents(self) -> int:
        """
        오른쪽 스프레드에서 기존 부양가족 행을 모두 삭제

        1행과 2행의 관계코드를 비교하여 다르면 2행을 F5+y로 삭제.
        같아질 때까지 반복.

        Returns:
            삭제한 행 수
        """
        deleted = 0
        max_attempts = 100  # 무한루프 방지

        for _ in range(max_attempts):
            if self.check_stop_key():
                break

            # 1행 관계코드 읽기
            self.input_handler.type_keys_with_delay(self.right_spread, "^{HOME}", pause=0.05)
            self.input_handler.type_keys_with_delay(self.right_spread, "{HOME}", pause=0.05)
            row1_value = self.input_handler.copy_from_control(self.right_spread)

            # 2행 관계코드 읽기
            self.input_handler.type_keys_with_delay(self.right_spread, "{DOWN}", pause=0.05)
            row2_value = self.input_handler.copy_from_control(self.right_spread)

            if row1_value == row2_value:
                # 같으면 부양가족 없음, 종료
                break

            # 다르면 2행 삭제 (F5 → 글로벌 y)
            self.input_handler.type_keys_with_delay(self.right_spread, "{F5}", sleep_after=0.2, pause=0.05)
            keyboard.press_and_release('y')
            time.sleep(0.1 * self.global_delay)
            deleted += 1
            self.log("DEBUG", f"  기존 부양가족 행 삭제 ({deleted}번째)")

        if deleted > 0:
            self.log("INFO", f"  → 기존 부양가족 {deleted}명 삭제 완료")

        return deleted

    def _process_with_employee_no(self, emp_no: str) -> Dict:
        """
        이미 읽은 사번으로 부양가족 처리

        Args:
            emp_no: 사번

        Returns:
            결과 딕셔너리
        """
        start_time = time.time()

        if not emp_no:
            return {'status': 'error', 'reason': 'empty_employee_no'}

        # CSV에서 사원 이름 가져오기 (로그용)
        emp_name = ""
        if emp_no in self.csv_data:
            first_dep = self.csv_data[emp_no][0]
            emp_name = first_dep.employee_name

        self.log("INFO", f"처리 시작: {emp_no} ({emp_name})")

        # 오른쪽 스프레드로 포커스 이동 → 기존 부양가족 삭제
        self.right_spread.set_focus()
        self.clear_existing_dependents()

        # CSV 데이터 찾기
        if emp_no not in self.csv_data:
            self.log("INFO", f"  → CSV 데이터 없음, 기존 삭제만 수행")
            self.left_spread.set_focus()
            time.sleep(0.1 * self.global_delay)
            return {
                'status': 'skip',
                'reason': 'no_csv_data',
                'employee_no': emp_no,
                'employee_name': emp_name
            }

        # 부양가족 필터링 (본인 제외) 및 정렬
        all_dependents = self.csv_data[emp_no]
        dependents = [d for d in all_dependents if d.relationship_code != '0']
        dependents = sorted(dependents, key=lambda d: (d.relationship_code, d.name))

        if not dependents:
            self.log("INFO", f"  → 부양가족 없음, 기존 삭제만 수행")
            self.left_spread.set_focus()
            time.sleep(0.1 * self.global_delay)
            return {
                'status': 'skip',
                'reason': 'no_dependents',
                'employee_no': emp_no,
                'employee_name': emp_name
            }

        # 입력 시작 위치로 이동
        self.input_handler.type_keys_with_delay(self.right_spread, "^{HOME}", pause=0.05)
        self.input_handler.type_keys_with_delay(self.right_spread, "{HOME}", pause=0.05)
        self.input_handler.type_keys_with_delay(self.right_spread, "{DOWN}", pause=0.05)

        # 각 부양가족 입력
        self.log("INFO", f"  → 부양가족 {len(dependents)}명 입력 시작")
        success_count = 0

        for dep in dependents:
            # 중지 요청 체크
            if self.check_stop_key():
                self.log("WARNING", "  → 중지 요청으로 부양가족 입력 중단")
                break

            if self.input_dependent(dep):
                success_count += 1
                self.log("SUCCESS", f"    [OK] {dep.name} (관계: {dep.relationship_code})")
            else:
                self.log("ERROR", f"    [X] {dep.name} 실패")

        self.log("INFO", f"  → 완료: {success_count}/{len(dependents)}")

        # 왼쪽 스프레드로 포커스 복귀
        self.left_spread.set_focus()
        time.sleep(0.1 * self.global_delay)

        elapsed = time.time() - start_time
        self.log("DEBUG", f"⏱️ 사원 1명 전체 처리: {elapsed:.2f}초")

        return {
            'status': 'success',
            'employee_no': emp_no,
            'employee_name': emp_name,
            'total': len(dependents),
            'success': success_count
        }

    def process_current_employee(self) -> Dict:
        """
        현재 사원의 부양가족 전체 처리

        Returns:
            결과 딕셔너리
        """
        try:
            emp_no = self.read_current_employee_no()
        except Exception as e:
            self.log("ERROR", f"사원 정보 읽기 실패: {e}")
            return {'status': 'error', 'reason': 'read_failed'}

        return self._process_with_employee_no(emp_no)

    def run(self, count: int = None, dry_run: bool = False) -> Dict:
        """
        전체 자동화 실행

        Args:
            count: 처리할 사원 수 (None이면 빈 칸까지 전체)
            dry_run: True면 실제 입력 안함 (테스트용)

        Returns:
            결과 요약 딕셔너리
        """
        start_time = time.time()

        self.log("INFO", "=== 부양가족 대량 입력 시작 ===")
        self.log("INFO", f"CSV 파일: {self.csv_reader.csv_path}")
        self.log("INFO", f"CSV 사원 수: {len(self.csv_data)}")
        if count is None:
            self.log("INFO", f"처리할 사원 수: 전체 (빈 칸까지)")
        else:
            self.log("INFO", f"처리할 사원 수: {count}명")

        if dry_run:
            self.log("INFO", "!!! DRY RUN 모드 - 실제 입력 안함 !!!")

        # 실행 전 스프레드 상태 재확인
        self.log("INFO", "화면 상태 확인 중...")
        self.spread_controller.verify_spreads()
        self.log("INFO", "✓ 화면 상태 정상")

        # 시작 위치로 이동
        if not self.start_from_current:
            self.log("INFO", "첫 번째 사원으로 이동...")
            self.input_handler.type_keys_with_delay(self.left_spread, "^{HOME}", pause=0.05)
            self.input_handler.type_keys_with_delay(self.left_spread, "{HOME}", pause=0.05)
            self.input_handler.type_keys_with_delay(self.left_spread, "{LEFT}", pause=0.05)
        else:
            self.log("INFO", "현재 위치에서 시작...")

        # 각 사원 처리
        results = []

        if count is None:
            # 전체 처리 모드: 빈 칸까지
            i = 0
            prev_emp_no = None
            while True:
                self.check_stop_key()

                if self.check_stop_key():
                    self.log("WARNING", "중지 요청으로 처리 중단")
                    break

                self.log("INFO", f"\n[{i+1}]")

                emp_no = self.read_current_employee_no()

                # 빈 칸 감지
                if i > 0 and emp_no == prev_emp_no:
                    self.log("INFO", "빈 칸 도달, 처리 종료")
                    break

                if not emp_no:
                    self.log("INFO", "빈 사번, 처리 종료")
                    break

                if not dry_run:
                    result = self._process_with_employee_no(emp_no)
                    results.append(result)
                else:
                    emp_name = ""
                    if emp_no in self.csv_data:
                        emp_name = self.csv_data[emp_no][0].employee_name
                    self.log("INFO", f"읽음: {emp_no} ({emp_name})")
                    has_data = emp_no in self.csv_data
                    self.log("INFO", f"  → CSV 데이터: {'있음' if has_data else '없음'}")

                prev_emp_no = emp_no
                self.input_handler.type_keys_with_delay(self.left_spread, "{DOWN}", pause=0.05)
                i += 1

        else:
            # 개수 지정 모드
            for i in range(count):
                self.check_stop_key()

                if self.check_stop_key():
                    self.log("WARNING", "중지 요청으로 처리 중단")
                    break

                self.log("INFO", f"\n[{i+1}/{count}]")

                if not dry_run:
                    result = self.process_current_employee()
                    results.append(result)
                else:
                    emp_no = self.read_current_employee_no()
                    emp_name = ""
                    if emp_no and emp_no in self.csv_data:
                        emp_name = self.csv_data[emp_no][0].employee_name
                    self.log("INFO", f"읽음: {emp_no} ({emp_name})")
                    has_data = emp_no in self.csv_data if emp_no else False
                    self.log("INFO", f"  → CSV 데이터: {'있음' if has_data else '없음'}")

                if i < count - 1:
                    self.input_handler.type_keys_with_delay(self.left_spread, "{DOWN}", pause=0.05)

        # 결과 집계
        elapsed_time = time.time() - start_time
        minutes = int(elapsed_time // 60)
        seconds = elapsed_time % 60

        if not dry_run:
            summary = self._summarize_results(results)
            self.log("INFO", "\n=== 결과 요약 ===")
            self.log("INFO", f"처리 사원: {summary['processed']}")
            self.log("INFO", f"성공: {summary['success']}")
            self.log("INFO", f"건너뜀: {summary['skipped']}")
            self.log("INFO", f"실패: {summary['failed']}")
            self.log("INFO", f"입력 부양가족: {summary['total_dependents']}")
            if minutes > 0:
                self.log("INFO", f"총 소요시간: {minutes}분 {seconds:.1f}초")
            else:
                self.log("INFO", f"총 소요시간: {seconds:.1f}초")

            self.cleanup()
            return summary
        else:
            self.log("INFO", "\n=== DRY RUN 완료 ===")
            if minutes > 0:
                self.log("INFO", f"총 소요시간: {minutes}분 {seconds:.1f}초")
            else:
                self.log("INFO", f"총 소요시간: {seconds:.1f}초")

            self.cleanup()
            return {}

    def _summarize_results(self, results: List[Dict]) -> Dict:
        """결과 리스트 집계"""
        summary = {
            'processed': len(results),
            'success': len([r for r in results if r['status'] == 'success']),
            'skipped': len([r for r in results if r['status'] == 'skip']),
            'failed': len([r for r in results if r['status'] == 'error']),
            'total_dependents': sum(r.get('success', 0) for r in results)
        }
        return summary


def main():
    """메인 함수"""
    parser = argparse.ArgumentParser(
        description='부양가족 대량 입력 자동화',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  # 드라이런 (처음 3명만 읽기 테스트)
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --count 3 --dry-run

  # 실제 실행 (처음 1명)
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --count 1

  # 실제 실행 (처음 5명)
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --count 5

  # 실제 실행 (전체)
  python bulk_dependent_input.py --csv "테스트 데이터.csv"
        """
    )

    parser.add_argument('--csv', required=True, help='CSV 파일 경로')
    parser.add_argument('--count', type=int, default=None, help='처리할 사원 수 (미지정 시 전체)')
    parser.add_argument('--dry-run', action='store_true', help='실제 입력 없이 테스트')
    parser.add_argument('--delay', type=float, default=1.0, help='전역 지연 시간 배율 (0.5~2.0, 기본값: 1.0)')

    args = parser.parse_args()

    # 실행
    bulk = None
    try:
        bulk = BulkDependentInput(args.csv, global_delay=args.delay)
        bulk.run(count=args.count, dry_run=args.dry_run)
    except KeyboardInterrupt:
        print("\n\n중단됨 (Ctrl+C)")
        if bulk:
            bulk.cleanup()
        sys.exit(1)
    except Exception as e:
        print(f"\n\n치명적 오류: {e}")
        import traceback
        traceback.print_exc()
        if bulk:
            bulk.cleanup()
        sys.exit(1)


if __name__ == "__main__":
    main()
