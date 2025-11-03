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
from typing import Dict, List, Tuple
import win32gui
import keyboard

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

    def __init__(self, csv_path: str, verbose: bool = False):
        """
        초기화 및 연결

        Args:
            csv_path: CSV 파일 경로
            verbose: True면 DEBUG 로그 출력, False면 숨김 (기본값: False)
        """
        self.verbose = verbose
        self.stop_requested = False  # 중지 플래그
        self.pause_press_count = 0  # Pause 키 연속 누름 횟수
        self.last_pause_time = 0  # 마지막 Pause 키 누름 시간

        # Pause 키 리스너 등록
        keyboard.on_press_key('pause', self._on_pause_press)

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
            print(f"  ⚠️  부양가족정보 탭이 선택되어 있는지 확인하세요!")
        except Exception as e:
            print(f"  [X] 연결 실패: {e}")
            print("    → 사원등록 프로그램이 실행 중인지 확인하세요")
            sys.exit(1)

        # 왼쪽/오른쪽 스프레드 찾기 및 포커스 (한 번만!)
        try:
            self.left_spread = self._find_left_spread()
            self.right_spread = self._find_right_spread()
            print(f"  [OK] 스프레드 찾기 완료 (왼쪽/오른쪽)")

            # 왼쪽 포커스 설정 (마우스 움직임을 최소화하기 위해 초기화 시 한 번만)
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

    def _find_right_spread(self):
        """
        오른쪽 부양가족 스프레드 찾기

        Returns:
            pywinauto wrapper object
        """
        spreads = self.dlg.children(class_name="fpUSpread80")
        if len(spreads) < 2:
            raise Exception("오른쪽 스프레드를 찾을 수 없습니다")

        # X 좌표로 정렬하여 가장 오른쪽 반환
        spreads.sort(key=lambda s: s.rectangle().left)
        return spreads[-1]

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

    def _on_pause_press(self, event):
        """
        Pause 키 눌림 콜백 (keyboard 패키지)

        Args:
            event: keyboard.KeyboardEvent
        """
        current_time = time.time()

        # 2초 이내에 이전 키 입력이 있었는지 확인
        if current_time - self.last_pause_time <= 2.0:
            self.pause_press_count += 1
        else:
            # 2초 넘었으면 카운트 리셋
            self.pause_press_count = 1

        self.last_pause_time = current_time

        # 3번 눌렸으면 중지
        if self.pause_press_count >= 3:
            self.log("WARNING", "Pause 키 3번 감지 - 중지 요청")
            self.stop_requested = True
            self.pause_press_count = 0  # 리셋
        else:
            self.log("INFO", f"Pause 키 감지 ({self.pause_press_count}/3)")

    def check_stop_key(self) -> bool:
        """
        중지 요청 상태 체크 (keyboard 콜백에서 설정됨)

        Returns:
            True면 중지 요청, False면 계속 진행
        """
        return self.stop_requested

    def stop(self):
        """
        자동화 중지 요청

        실행 중인 run() 메서드는 다음 체크 포인트에서 중지됩니다.
        """
        self.log("WARNING", "중지 요청됨 - 현재 작업 완료 후 중단...")
        self.stop_requested = True

    def cleanup(self):
        """
        리소스 정리 (keyboard 후크 해제)
        """
        try:
            keyboard.unhook_key('pause')
            self.log("INFO", "Pause 키 리스너 해제됨")
        except Exception as e:
            self.log("WARNING", f"Pause 키 리스너 해제 실패: {e}")

    def _type_keys(self, control, keys: str, **kwargs):
        """
        디버그용 type_keys 래퍼

        Args:
            control: pywinauto 컨트롤 (left_spread 또는 right_spread)
            keys: 입력할 키
            **kwargs: type_keys에 전달할 추가 인자
        """
        # 어느 스프레드인지 구분
        control_name = "LEFT " if control == self.left_spread else "RIGHT"
        self.log("DEBUG", f"[{control_name}] type_keys: '{keys}'")
        control.type_keys(keys, **kwargs)

    def _type_keys_with_delay(self, control, keys: str, sleep_after: float = 0.1, **kwargs):
        """
        type_keys + sleep을 통합한 메서드

        Args:
            control: pywinauto 컨트롤 (left_spread 또는 right_spread)
            keys: 입력할 키
            sleep_after: 키 입력 후 대기 시간 (기본 0.1초)
            **kwargs: type_keys에 전달할 추가 인자
        """
        self._type_keys(control, keys, **kwargs)
        time.sleep(sleep_after)

    def _send_copy(self):
        """
        클립보드 복사 (창 활성화 방식)

        fpUSpread80은 SendMessage/COM으로 접근 불가능.
        type_keys는 전역 키보드 입력이므로 활성 창에 입력됨.

        해결책: 사원등록 창을 잠깐 활성화하고 복사 후 원래 창으로 복귀
        """
        self._type_keys_with_delay(self.left_spread, "^c", pause=0.05)

    def _paste_text(self, control, text: str, sleep_after: float = 0.15):
        """
        클립보드 복사 후 붙여넣기 (긴 텍스트 입력 최적화)

        type_keys 방식보다 빠름:
        - 한글 성명: 0.75초 → 0.15초
        - 13자리 번호: 0.93초 → 0.15초

        Args:
            control: pywinauto 컨트롤
            text: 입력할 텍스트
            sleep_after: 붙여넣기 후 대기 시간 (기본 0.1초)
        """
        # 클립보드에 텍스트 복사
        pyperclip.copy(text)
        # Ctrl+V로 붙여넣기
        self._type_keys_with_delay(control, "^v", sleep_after=sleep_after, pause=0.05)

    def read_current_employee_no(self) -> str:
        """
        현재 선택된 사원의 사번 읽기 (클립보드 방식)

        Returns:
            사번
        """
        start_time = time.time()

        # HOME → RIGHT → Ctrl+C → 사번
        self._type_keys_with_delay(self.left_spread, "{HOME}", pause=0.05)
        self._type_keys_with_delay(self.left_spread, "{RIGHT}", pause=0.05)  # 체크박스 건너뛰기
        self._send_copy()

        employee_no = pyperclip.paste().strip()

        elapsed = time.time() - start_time
        self.log("DEBUG", f"⏱️ 사번 읽기: {elapsed:.2f}초")

        return employee_no

    def input_dependent(self, dep: DependentData) -> bool:
        """
        부양가족 한 명의 데이터 입력 (type_keys 방식)

        Args:
            dep: 부양가족 데이터

        Returns:
            성공 여부
        """
        start_time = time.time()

        try:
            # 1. 관계코드 입력 (자동으로 다음 열로 이동)
            t1 = time.time()
            self._type_keys_with_delay(self.right_spread, dep.relationship_code, with_spaces=False, pause=0.05)
            self.log("DEBUG", f"  ⏱️ 관계코드: {time.time()-t1:.2f}초")

            # 2. 성명 입력
            t2 = time.time()
            self._type_keys_with_delay(self.right_spread, dep.name, with_spaces=False, pause=0.05)

            # ENTER (다음 열로 이동) - 대기 시간 충분히 증가
            self._type_keys_with_delay(self.right_spread, "{ENTER}", pause=0.05)
            self.log("DEBUG", f"  ⏱️ 성명: {time.time()-t2:.2f}초")

            # 3. 내/외국인 입력 (N=1, Y=2 변환, 자동으로 다음 열로 이동)
            t3 = time.time()
            nationality_code = self._convert_nationality(dep.nationality)
            self._type_keys_with_delay(self.right_spread, nationality_code, with_spaces=False, pause=0.05)
            self.log("DEBUG", f"  ⏱️ 내/외국인: {time.time()-t3:.2f}초")

            # 4. 번호 타입 입력 (1=주민, 2=외국인, 3=여권, 자동으로 다음 열로 이동)
            t4 = time.time()
            id_type_code, id_number_clean = self._clean_id_number(dep.id_number)
            self._type_keys_with_delay(self.right_spread, id_type_code, with_spaces=False, pause=0.05)
            self.log("DEBUG", f"  ⏱️ 번호타입: {time.time()-t4:.2f}초")

            # 5. 주민등록번호/외국인등록번호/여권번호 입력 (자동으로 다음 열로 이동)
            t5 = time.time()
            self._type_keys_with_delay(self.right_spread, id_number_clean, with_spaces=False, pause=0.05)
            self.log("DEBUG", f"  ⏱️ 번호입력: {time.time()-t5:.2f}초")

            # 6. 나이는 자동 입력되므로 건너뜀

            # 7. 기본공제여부 (Y=자동입력, N=0 입력 필요, 자동으로 다음 열로 이동)
            t7 = time.time()
            if dep.basic_deduction.strip().upper() == 'N':
                self._type_keys_with_delay(self.right_spread, "0", with_spaces=False, pause=0.05)
            # Y인 경우 자동으로 입력되므로 아무것도 안함
            else:
                self._type_keys_with_delay(self.right_spread, "{RIGHT}", with_spaces=False, pause=0.05)
            self.log("DEBUG", f"  ⏱️ 기본공제: {time.time()-t7:.2f}초")

            # 8. 자녀공제 (연말관계가 4이고, 기본공제가 Y인 경우만)
            t8 = time.time()
            if dep.relationship_code == '4' and dep.basic_deduction.strip().upper() == 'Y':
                # 장애인 열 건너뛰기 (RIGHT)
                self._type_keys_with_delay(self.right_spread, "{RIGHT}", pause=0.05)

                # 자녀공제 입력 (N=0, Y=1, 자동으로 다음 열로 이동)
                child_code = '1' if dep.child_deduction.strip().upper() == 'Y' else '0'
                self._type_keys_with_delay(self.right_spread, child_code, with_spaces=False, pause=0.05)
            self.log("DEBUG", f"  ⏱️ 자녀공제: {time.time()-t8:.2f}초")

            # 다음 부양가족으로 이동 (DOWN + HOME)
            t9 = time.time()
            self._type_keys_with_delay(self.right_spread, "{HOME}", pause=0.05)
            self._type_keys_with_delay(self.right_spread, "{DOWN}", pause=0.05)
            self.log("DEBUG", f"  ⏱️ 다음행이동: {time.time()-t9:.2f}초")

            elapsed = time.time() - start_time
            self.log("DEBUG", f"⏱️ 부양가족 1명 전체: {elapsed:.2f}초")

            return True

        except Exception as e:
            self.log("ERROR", f"input_dependent 실패: {e}")
            return False

    def _convert_nationality(self, nationality: str) -> str:
        """
        내/외국인 코드 변환

        Args:
            nationality: CSV의 내/외국인 값 (N, Y, 내, 외 등)

        Returns:
            변환된 코드 (1=내국인, 2=외국인)
        """
        # 대소문자 무시하고 변환
        val = nationality.strip().upper()

        # 내국인: N, 내, 1
        if val in ['N', '내', '내국인', '1']:
            return '1'
        # 외국인: Y, 외, 2
        elif val in ['Y', '외', '외국인', '2']:
            return '2'
        else:
            # 기본값: 내국인
            self.log("WARNING", f"알 수 없는 내/외국인 값: '{nationality}' → 내국인(1)로 처리")
            return '1'

    def _clean_id_number(self, id_number: str) -> Tuple[str, str]:
        """
        주민등록번호/외국인등록번호/여권번호 정제 및 타입 판별

        Args:
            id_number: 원본 번호

        Returns:
            (타입_코드, 정제된_번호) 튜플
            - 타입_코드: '1'=주민등록번호, '2'=외국인등록번호, '3'=여권번호
            - 정제된_번호: 정제된 번호 문자열
        """
        import re

        # 1. 여권번호 판별 (알파벳 포함)
        if re.search(r'[A-Za-z]', id_number):
            # 여권번호: 공백, 하이픈 제거, 대문자 변환
            cleaned = re.sub(r'[\s\-]', '', id_number).upper()
            self.log("DEBUG", f"여권번호(3) 감지: {id_number} → {cleaned}")
            return ('3', cleaned)

        # 2. 주민등록번호 또는 외국인등록번호 (13자리 숫자)
        # 하이픈 제거
        cleaned = id_number.replace("-", "").strip()

        # 숫자만 추출
        cleaned = re.sub(r'[^0-9]', '', cleaned)

        # 13자리 검증
        if len(cleaned) == 13:
            # 7번째 자리로 내외국인 구분
            gender_code = cleaned[6]
            if gender_code in ['5', '6', '7', '8']:
                self.log("DEBUG", f"외국인등록번호(2) 감지: {id_number} → {cleaned}")
                return ('2', cleaned)
            else:
                self.log("DEBUG", f"주민등록번호(1): {id_number} → {cleaned}")
                return ('1', cleaned)
        else:
            self.log("WARNING", f"비정상 번호 길이 ({len(cleaned)}자리): {id_number}")
            # 기본값: 주민등록번호로 처리
            return ('1', cleaned)

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

        # CSV 데이터 찾기
        if emp_no not in self.csv_data:
            self.log("INFO", f"  → CSV 데이터 없음, 건너뜀")
            return {
                'status': 'skip',
                'reason': 'no_csv_data',
                'employee_no': emp_no,
                'employee_name': emp_name
            }

        # 부양가족 필터링 (본인 제외) 및 정렬
        all_dependents = self.csv_data[emp_no]
        dependents = [d for d in all_dependents if d.relationship_code != '0']

        # (연말관계, 이름) 순서로 정렬
        dependents = sorted(dependents, key=lambda d: (d.relationship_code, d.name))

        if not dependents:
            self.log("INFO", f"  → 부양가족 없음, 건너뜀")
            return {
                'status': 'skip',
                'reason': 'no_dependents',
                'employee_no': emp_no,
                'employee_name': emp_name
            }

        # 오른쪽 스프레드로 포커스 이동
        self.right_spread.set_focus()
        self._type_keys_with_delay(self.right_spread, "^{HOME}", pause=0.05)
        self._type_keys_with_delay(self.right_spread, "{HOME}", pause=0.05)

        # 첫 번째 행(본인)은 건너뛰고 두 번째 행으로 이동
        self._type_keys_with_delay(self.right_spread, "{DOWN}", pause=0.05)

        # 각 부양가족 입력
        self.log("INFO", f"  → 부양가족 {len(dependents)}명 입력 시작")
        success_count = 0

        for dep in dependents:
            # 중지 요청 체크
            if self.stop_requested:
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
        time.sleep(0.1)

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
            {
                'status': 'success' | 'skip' | 'error',
                'employee_no': str,
                'employee_name': str,
                'total': int,
                'success': int,
                'reason': str  # skip/error인 경우
            }
        """
        # 사번 읽기
        try:
            emp_no = self.read_current_employee_no()
        except Exception as e:
            self.log("ERROR", f"사원 정보 읽기 실패: {e}")
            return {'status': 'error', 'reason': 'read_failed'}

        # 이미 읽은 사번으로 처리
        return self._process_with_employee_no(emp_no)

    def run(self,
            count: int = None,
            dry_run: bool = False) -> Dict:
        """
        전체 자동화 실행

        Args:
            count: 처리할 사원 수 (None이면 빈 칸까지 전체)
            dry_run: True면 실제 입력 안함 (테스트용)

        Returns:
            결과 요약 딕셔너리
        """
        # 시작 시간 기록
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

        # 1. 시작 위치로 이동 (Ctrl+Home)
        self.log("INFO", "첫 번째 사원으로 이동...")
        self._type_keys_with_delay(self.left_spread, "^{HOME}", pause=0.05)
        self._type_keys_with_delay(self.left_spread, "{HOME}", pause=0.05)
        self._type_keys_with_delay(self.left_spread, "{LEFT}", pause=0.05)

        # 2. 각 사원 처리
        results = []

        if count is None:
            # 전체 처리 모드: 빈 칸까지
            i = 0
            prev_emp_no = None
            while True:
                # F12 키 체크
                self.check_stop_key()

                # 중지 요청 체크
                if self.stop_requested:
                    self.log("WARNING", "중지 요청으로 처리 중단")
                    break

                self.log("INFO", f"\n[{i+1}]")

                # 사번 먼저 읽어서 빈 칸인지 확인
                emp_no = self.read_current_employee_no()

                # 빈 칸 감지: 이전 사번과 같으면 빈 칸 (클립보드에 이전 값이 남아있음)
                if i > 0 and emp_no == prev_emp_no:
                    self.log("INFO", "빈 칸 도달, 처리 종료")
                    break

                if not emp_no:
                    self.log("INFO", "빈 사번, 처리 종료")
                    break

                if not dry_run:
                    # 이미 읽은 사번으로 처리
                    result = self._process_with_employee_no(emp_no)
                    results.append(result)
                else:
                    # dry run: 읽기만
                    emp_name = ""
                    if emp_no in self.csv_data:
                        emp_name = self.csv_data[emp_no][0].employee_name
                    self.log("INFO", f"읽음: {emp_no} ({emp_name})")
                    has_data = emp_no in self.csv_data
                    self.log("INFO", f"  → CSV 데이터: {'있음' if has_data else '없음'}")

                # 다음 사원으로
                prev_emp_no = emp_no  # 현재 사번 저장
                self._type_keys_with_delay(self.left_spread, "{DOWN}", pause=0.05)
                i += 1

        else:
            # 개수 지정 모드
            for i in range(count):
                # F12 키 체크
                self.check_stop_key()

                # 중지 요청 체크
                if self.stop_requested:
                    self.log("WARNING", "중지 요청으로 처리 중단")
                    break

                self.log("INFO", f"\n[{i+1}/{count}]")

                if not dry_run:
                    result = self.process_current_employee()
                    results.append(result)
                else:
                    # dry run: 읽기만
                    emp_no = self.read_current_employee_no()
                    emp_name = ""
                    if emp_no and emp_no in self.csv_data:
                        emp_name = self.csv_data[emp_no][0].employee_name
                    self.log("INFO", f"읽음: {emp_no} ({emp_name})")
                    has_data = emp_no in self.csv_data if emp_no else False
                    self.log("INFO", f"  → CSV 데이터: {'있음' if has_data else '없음'}")

                # 다음 사원으로
                if i < count - 1:
                    self._type_keys_with_delay(self.left_spread, "{DOWN}", pause=0.05)

        # 3. 결과 집계
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

            # 리소스 정리
            self.cleanup()
            return summary
        else:
            self.log("INFO", "\n=== DRY RUN 완료 ===")
            if minutes > 0:
                self.log("INFO", f"총 소요시간: {minutes}분 {seconds:.1f}초")
            else:
                self.log("INFO", f"총 소요시간: {seconds:.1f}초")

            # 리소스 정리
            self.cleanup()
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
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --count 3 --dry-run

  # 실제 실행 (처음 1명)
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --count 1

  # 실제 실행 (처음 5명)
  python bulk_dependent_input.py --csv "테스트 데이터.csv" --count 5

  # 실제 실행 (전체)
  python bulk_dependent_input.py --csv "테스트 데이터.csv"
        """
    )

    parser.add_argument(
        '--csv',
        required=True,
        help='CSV 파일 경로'
    )
    parser.add_argument(
        '--count',
        type=int,
        default=None,
        help='처리할 사원 수 (미지정 시 전체)'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='실제 입력 없이 테스트'
    )

    args = parser.parse_args()

    # 실행
    bulk = None
    try:
        bulk = BulkDependentInput(args.csv)

        bulk.run(
            count=args.count,
            dry_run=args.dry_run
        )
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
