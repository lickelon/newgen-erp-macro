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
from typing import Optional, Dict, List, Tuple

try:
    import pyperclip
except ImportError:
    print("ERROR: pyperclip이 설치되지 않았습니다.")
    print("설치: uv add pyperclip")
    sys.exit(1)

from pywinauto import Application
from src.csv_reader import CSVReader, DependentData


class BulkDependentInput:
    """부양가족 대량 입력 자동화"""

    def __init__(self, csv_path: str):
        """
        초기화 및 연결

        Args:
            csv_path: CSV 파일 경로
        """
        print(f"초기화 중...")

        # CSV 데이터 로드
        self.csv_reader = CSVReader(csv_path)
        self.csv_reader.read()
        self.csv_data = self.csv_reader.group_by_employee()
        print(f"  ✓ CSV 로드: {len(self.csv_data)}명 사원")

        # pywinauto 연결
        try:
            self.app = Application(backend='win32')
            self.dlg = self.app.connect(title_re=".*사원등록.*")
            print(f"  ✓ 사원등록 프로그램 연결")
        except Exception as e:
            print(f"  ✗ 연결 실패: {e}")
            print("    → 사원등록 프로그램이 실행 중인지 확인하세요")
            sys.exit(1)

        # 왼쪽 스프레드 찾기
        try:
            self.left_spread = self._find_left_spread()
            print(f"  ✓ 왼쪽 사원 목록 찾기 완료")
        except Exception as e:
            print(f"  ✗ 스프레드 찾기 실패: {e}")
            sys.exit(1)

        # 로그 설정
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = log_dir / f"bulk_input_{timestamp}.log"
        print(f"  ✓ 로그 파일: {self.log_file}")

    def _find_left_spread(self):
        """
        왼쪽 사원 목록 스프레드 찾기

        Returns:
            pywinauto wrapper object
        """
        spreads = self.dlg.children(class_name="fpUSpread80")
        if not spreads:
            raise Exception("fpUSpread80 컨트롤을 찾을 수 없습니다")

        # X 좌표로 정렬하여 가장 왼쪽 반환
        spreads.sort(key=lambda s: s.rectangle().left)
        return spreads[0]

    def log(self, level: str, message: str):
        """
        로그 기록

        Args:
            level: 로그 레벨 (INFO, SUCCESS, WARNING, ERROR)
            message: 메시지
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] {level}: {message}"

        # 콘솔 출력
        print(log_line)

        # 파일 기록
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(log_line + '\n')

    def count_employees(self) -> int:
        """
        왼쪽 목록의 총 사원 수 확인

        Returns:
            사원 수
        """
        # TODO: 구현 필요
        # 1. 왼쪽 스프레드 포커스
        # 2. HOME 키로 처음으로
        # 3. 루프: Ctrl+C로 복사 → 빈 값 만날 때까지 DOWN
        # 4. HOME으로 다시 처음으로
        self.log("WARNING", "count_employees() 미구현 - 임시로 0 반환")
        return 0

    def read_current_employee(self) -> Tuple[str, str]:
        """
        현재 선택된 사원의 사번과 이름 읽기 (클립보드 방식)

        Returns:
            (사번, 이름) 튜플
        """
        # TODO: 구현 필요
        # 1. 포커스 확인
        # 2. HOME → Ctrl+C → 사번
        # 3. RIGHT → Ctrl+C → 이름
        # 4. HOME으로 다시 처음
        self.log("WARNING", "read_current_employee() 미구현")
        return ("", "")

    def input_dependent(self, dep: DependentData) -> bool:
        """
        부양가족 한 명의 데이터 입력

        Args:
            dep: 부양가족 데이터

        Returns:
            성공 여부
        """
        # TODO: 구현 필요
        # 1. 관계코드 → TAB
        # 2. 이름 → TAB
        # 3. 내/외국인 → TAB
        # 4. 주민등록번호 → ENTER
        self.log("WARNING", f"input_dependent({dep.name}) 미구현")
        return False

    def process_current_employee(self) -> Dict:
        """
        현재 사원의 부양가족 전체 처리

        Returns:
            결과 딕셔너리
            {
                'status': 'success' | 'skip' | 'error',
                'employee_no': str,
                'employee_name': str,
                'total': int,
                'success': int,
                'reason': str  # skip/error인 경우
            }
        """
        # 1. 사번/이름 읽기
        try:
            emp_no, emp_name = self.read_current_employee()
        except Exception as e:
            self.log("ERROR", f"사원 정보 읽기 실패: {e}")
            return {'status': 'error', 'reason': 'read_failed'}

        if not emp_no:
            return {'status': 'error', 'reason': 'empty_employee_no'}

        self.log("INFO", f"처리 시작: {emp_no} ({emp_name})")

        # 2. CSV 데이터 찾기
        if emp_no not in self.csv_data:
            self.log("INFO", f"  → CSV 데이터 없음, 건너뜀")
            return {
                'status': 'skip',
                'reason': 'no_csv_data',
                'employee_no': emp_no,
                'employee_name': emp_name
            }

        # 3. 부양가족 필터링 (본인 제외)
        all_dependents = self.csv_data[emp_no]
        dependents = [d for d in all_dependents if d.relationship_code != '0']

        if not dependents:
            self.log("INFO", f"  → 부양가족 없음, 건너뜀")
            return {
                'status': 'skip',
                'reason': 'no_dependents',
                'employee_no': emp_no,
                'employee_name': emp_name
            }

        # 4. 각 부양가족 입력
        self.log("INFO", f"  → 부양가족 {len(dependents)}명 입력 시작")
        success_count = 0

        for dep in dependents:
            if self.input_dependent(dep):
                success_count += 1
                self.log("SUCCESS", f"    ✓ {dep.name} (관계: {dep.relationship_code})")
            else:
                self.log("ERROR", f"    ✗ {dep.name} 실패")

        self.log("INFO", f"  → 완료: {success_count}/{len(dependents)}")

        return {
            'status': 'success',
            'employee_no': emp_no,
            'employee_name': emp_name,
            'total': len(dependents),
            'success': success_count
        }

    def run(self,
            limit: Optional[int] = None,
            skip: int = 0,
            dry_run: bool = False) -> Dict:
        """
        전체 자동화 실행

        Args:
            limit: 처리할 최대 사원 수 (None = 전체)
            skip: 처음 건너뛸 사원 수
            dry_run: True면 실제 입력 안함 (테스트용)

        Returns:
            결과 요약 딕셔너리
        """
        self.log("INFO", "=== 부양가족 대량 입력 시작 ===")
        self.log("INFO", f"CSV 파일: {self.csv_reader.csv_path}")
        self.log("INFO", f"CSV 사원 수: {len(self.csv_data)}")

        if dry_run:
            self.log("INFO", "!!! DRY RUN 모드 - 실제 입력 안함 !!!")

        # 1. 왼쪽 목록 포커스
        self.log("INFO", "왼쪽 사원 목록 포커스...")
        self.left_spread.set_focus()
        time.sleep(0.5)

        # 2. 사원 수 확인
        self.log("INFO", "사원 수 확인 중...")
        total_employees = self.count_employees()
        self.log("INFO", f"왼쪽 목록 사원 수: {total_employees}")

        if total_employees == 0:
            self.log("WARNING", "사원이 없습니다. count_employees() 구현 필요")
            return {}

        # 3. 처리 범위 설정
        start_idx = skip
        end_idx = min(total_employees, skip + limit) if limit else total_employees

        self.log("INFO", f"처리 범위: {start_idx+1}번째 ~ {end_idx}번째 사원")

        # 4. 시작 위치로 이동
        self.log("INFO", "시작 위치로 이동...")
        self.dlg.type_keys("{HOME}", pause=0.1)
        time.sleep(0.3)

        for _ in range(skip):
            self.dlg.type_keys("{DOWN}", pause=0.1)
            time.sleep(0.1)

        # 5. 각 사원 처리
        results = []

        for i in range(start_idx, end_idx):
            self.log("INFO", f"\n[{i+1}/{total_employees}]")

            if not dry_run:
                result = self.process_current_employee()
                results.append(result)
            else:
                # dry run: 읽기만
                emp_no, emp_name = self.read_current_employee()
                self.log("INFO", f"읽음: {emp_no} ({emp_name})")
                has_data = emp_no in self.csv_data if emp_no else False
                self.log("INFO", f"  → CSV 데이터: {'있음' if has_data else '없음'}")

            # 다음 사원으로
            if i < end_idx - 1:
                self.dlg.type_keys("{DOWN}", pause=0.1)
                time.sleep(0.3)

        # 6. 결과 집계
        if not dry_run:
            summary = self._summarize_results(results)
            self.log("INFO", "\n=== 결과 요약 ===")
            self.log("INFO", f"처리 사원: {summary['processed']}")
            self.log("INFO", f"성공: {summary['success']}")
            self.log("INFO", f"건너뜀: {summary['skipped']}")
            self.log("INFO", f"실패: {summary['failed']}")
            self.log("INFO", f"입력 부양가족: {summary['total_dependents']}")

            return summary
        else:
            self.log("INFO", "\n=== DRY RUN 완료 ===")
            return {}

    def _summarize_results(self, results: List[Dict]) -> Dict:
        """
        결과 리스트 집계

        Args:
            results: process_current_employee() 결과 리스트

        Returns:
            요약 딕셔너리
        """
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
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 3 --dry-run

  # 실제 실행 (처음 5명)
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 5

  # 전체 실행
  python bulk_dependent_input.py --csv "테스트 데이터.csv"

  # 10번째부터 5명
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --skip 10 --limit 5
        """
    )

    parser.add_argument(
        '--csv',
        required=True,
        help='CSV 파일 경로'
    )
    parser.add_argument(
        '--limit',
        type=int,
        help='처리할 사원 수 제한'
    )
    parser.add_argument(
        '--skip',
        type=int,
        default=0,
        help='건너뛸 사원 수 (기본: 0)'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='실제 입력 없이 테스트'
    )

    args = parser.parse_args()

    # 실행
    try:
        bulk = BulkDependentInput(args.csv)
        bulk.run(
            limit=args.limit,
            skip=args.skip,
            dry_run=args.dry_run
        )
    except KeyboardInterrupt:
        print("\n\n중단됨 (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n치명적 오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
