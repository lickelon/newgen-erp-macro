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
import win32api
import win32con

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
        print(f"  [OK] CSV 로드: {len(self.csv_data)}명 사원")

        # pywinauto 연결
        try:
            self.app = Application(backend='win32')
            self.app.connect(title="사원등록")
            self.dlg = self.app.window(title="사원등록")
            print(f"  [OK] 사원등록 프로그램 연결")
        except Exception as e:
            print(f"  [X] 연결 실패: {e}")
            print("    → 사원등록 프로그램이 실행 중인지 확인하세요")
            sys.exit(1)

        # 왼쪽 스프레드 찾기 및 포커스 (한 번만!)
        try:
            self.left_spread = self._find_left_spread()
            print(f"  [OK] 왼쪽 사원 목록 찾기 완료")

            # 포커스 설정 (마우스 움직임을 최소화하기 위해 초기화 시 한 번만)
            self.left_spread.set_focus()
            time.sleep(0.3)
            print(f"  [OK] 왼쪽 사원 목록 포커스 설정")
        except Exception as e:
            print(f"  [X] 스프레드 찾기 실패: {e}")
            sys.exit(1)

        # 로그 설정
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = log_dir / f"bulk_input_{timestamp}.log"
        print(f"  [OK] 로그 파일: {self.log_file}")

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

    def _make_lparam(self, scan_code: int, extended: bool = False, prev_state: bool = False, transition: bool = False):
        """
        WM_KEYDOWN/WM_KEYUP의 lParam 생성

        Args:
            scan_code: 스캔 코드
            extended: 확장 키 여부 (Home, End, 화살표 등)
            prev_state: 이전 키 상태 (WM_KEYUP에서 True)
            transition: 전환 상태 (WM_KEYUP에서 True)

        Returns:
            lParam 값
        """
        lparam = 1  # 반복 횟수 = 1
        lparam |= (scan_code << 16)  # 스캔 코드
        if extended:
            lparam |= (1 << 24)  # 확장 키 플래그
        if prev_state:
            lparam |= (1 << 30)  # 이전 키 상태
        if transition:
            lparam |= (1 << 31)  # 전환 상태
        return lparam

    def _send_key(self, vk_code: int, scan_code: int, extended: bool = False, delay: float = 0.02):
        """
        왼쪽 스프레드에 키 전송 (SendMessage 사용, 포커스 불필요)

        Args:
            vk_code: 가상 키 코드 (VK_DOWN, VK_HOME 등)
            scan_code: 스캔 코드
            extended: 확장 키 여부
            delay: KEYDOWN과 KEYUP 사이의 대기 시간 (초)
        """
        hwnd = self.left_spread.handle

        # KEYDOWN
        lparam_down = self._make_lparam(scan_code, extended=extended)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lparam_down)
        time.sleep(delay)

        # KEYUP
        lparam_up = self._make_lparam(scan_code, extended=extended, prev_state=True, transition=True)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, lparam_up)

    def _send_ctrl_key(self, vk_code: int, scan_code: int, extended: bool = False):
        """
        왼쪽 스프레드에 Ctrl 조합키 전송

        Args:
            vk_code: 가상 키 코드 (VK_HOME, ord('C') 등)
            scan_code: 스캔 코드
            extended: 확장 키 여부
        """
        hwnd = self.left_spread.handle

        # Ctrl 누름
        lparam_ctrl_down = self._make_lparam(0x1D, extended=False)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, lparam_ctrl_down)
        time.sleep(0.02)

        # 키 누름
        lparam_key_down = self._make_lparam(scan_code, extended=extended)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lparam_key_down)
        time.sleep(0.02)

        # 키 뗌
        lparam_key_up = self._make_lparam(scan_code, extended=extended, prev_state=True, transition=True)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, lparam_key_up)
        time.sleep(0.02)

        # Ctrl 뗌
        lparam_ctrl_up = self._make_lparam(0x1D, extended=False, prev_state=True, transition=True)
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, lparam_ctrl_up)

    # 편의 메서드들
    def _send_home(self):
        """HOME 키 전송"""
        self._send_key(win32con.VK_HOME, 0x47, extended=True)

    def _send_down(self):
        """DOWN 키 전송"""
        self._send_key(win32con.VK_DOWN, 0x50, extended=True)

    def _send_right(self):
        """RIGHT 키 전송"""
        self._send_key(win32con.VK_RIGHT, 0x4D, extended=True)

    def _send_ctrl_home(self):
        """Ctrl+Home 전송"""
        self._send_ctrl_key(win32con.VK_HOME, 0x47, extended=True)

    def _send_copy(self):
        """
        클립보드 복사

        fpUSpread80은 SendMessage로 복사 불가능.
        문서(successful-method.md)에 따르면 dlg.type_keys()만 작동.
        """
        # 방법 1: 스프레드에 직접 전송 시도
        self.left_spread.type_keys("^c", pause=0.05, set_foreground=False)
        time.sleep(0.1)

    def count_employees(self) -> int:
        """
        왼쪽 목록의 총 사원 수 확인

        Returns:
            사원 수
        """
        # 1. Ctrl+Home으로 첫 번째 셀로 이동
        self._send_ctrl_home()
        time.sleep(0.2)

        # 2. RIGHT로 사번 칸으로 이동 (첫 번째 칸은 체크박스)
        self._send_right()
        time.sleep(0.2)

        # 3. 루프: Ctrl+C로 복사 → 이전과 같은 값이면 중단
        count = 0
        max_iterations = 1000  # 무한루프 방지
        prev_value = None
        same_count = 0
        max_same = 2  # 같은 값이 2번 연속이면 끝에 도달

        for _ in range(max_iterations):
            # Ctrl+C로 현재 셀 복사
            self._send_copy()
            time.sleep(0.2)

            # 클립보드 읽기
            value = pyperclip.paste().strip()

            # 빈 값이면 중단
            if not value:
                break

            # 이전 값과 같으면 카운트 증가 (끝에 도달한 것으로 판단)
            if value == prev_value:
                same_count += 1
                if same_count >= max_same:
                    break
            else:
                same_count = 0

            prev_value = value
            count += 1

            # DOWN 키로 다음 행
            self._send_down()
            time.sleep(0.1)

        # 4. Ctrl+Home으로 다시 처음으로
        self._send_ctrl_home()
        time.sleep(0.2)

        return count

    def read_current_employee(self) -> Tuple[str, str]:
        """
        현재 선택된 사원의 사번과 이름 읽기 (클립보드 방식)

        Returns:
            (사번, 이름) 튜플
        """
        # 1. HOME → RIGHT → Ctrl+C → 사번
        self._send_home()
        time.sleep(0.2)

        self._send_right()  # 체크박스 건너뛰기
        time.sleep(0.2)

        self._send_copy()
        time.sleep(0.2)

        employee_no = pyperclip.paste().strip()

        # 2. RIGHT → Ctrl+C → 이름
        self._send_right()
        time.sleep(0.2)

        self._send_copy()
        time.sleep(0.2)

        employee_name = pyperclip.paste().strip()

        # 3. HOME으로 다시 처음
        self._send_home()
        time.sleep(0.2)

        return (employee_no, employee_name)

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
                self.log("SUCCESS", f"    [OK] {dep.name} (관계: {dep.relationship_code})")
            else:
                self.log("ERROR", f"    [X] {dep.name} 실패")

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

        # 1. 사원 수 확인 (포커스는 이미 초기화 시 설정됨)
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
        self._send_ctrl_home()
        time.sleep(0.3)

        for _ in range(skip):
            self._send_down()
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
                self._send_down()
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
