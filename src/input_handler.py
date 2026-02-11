"""
입력 핸들러 모듈

키보드 입력 및 클립보드 처리 기능을 제공합니다.
"""

import time
import pyperclip
from typing import Tuple


class InputHandler:
    """입력 처리 클래스"""

    def __init__(self, left_spread, right_spread, global_delay: float = 1.0, log_callback=None):
        """
        초기화

        Args:
            left_spread: 왼쪽 스프레드 컨트롤
            right_spread: 오른쪽 스프레드 컨트롤
            global_delay: 전역 지연 시간 배율 (0.5~2.0)
            log_callback: 로그 출력 콜백 함수 (level, message)
        """
        self.left_spread = left_spread
        self.right_spread = right_spread
        self.global_delay = max(0.5, min(2.0, global_delay))
        self.log_callback = log_callback

    def _log(self, level: str, message: str):
        """로그 출력 (콜백이 있으면 사용)"""
        if self.log_callback:
            self.log_callback(level, message)

    def type_keys(self, control, keys: str, **kwargs):
        """
        키보드 입력 (디버그 로그 포함)

        Args:
            control: pywinauto 컨트롤
            keys: 입력할 키
            **kwargs: type_keys에 전달할 추가 인자

        Raises:
            Exception: ElementNotVisible 등 입력 실패 시
        """
        # 어느 스프레드인지 구분
        control_name = "LEFT " if control == self.left_spread else "RIGHT"
        self._log("DEBUG", f"[{control_name}] type_keys: '{keys}'")

        try:
            control.type_keys(keys, **kwargs)
        except Exception as e:
            # ElementNotVisible 예외를 명확한 메시지로 변환
            if "ElementNotVisible" in str(type(e).__name__):
                error_msg = "❌ 부양가족상세 탭이 열려있지 않습니다!\n\n"
                error_msg += "다음을 확인하세요:\n"
                error_msg += "1. 사원등록 창에서 '부양가족정보' 탭을 선택했는지\n"
                error_msg += "2. 오른쪽 부양가족상세 영역이 표시되어 있는지\n"
                error_msg += "3. 실행 중에 다른 탭으로 이동하지 않았는지"
                raise Exception(error_msg)
            else:
                # 다른 예외는 그대로 전달
                raise

    def type_keys_with_delay(self, control, keys: str, sleep_after: float = 0.1, **kwargs):
        """
        키보드 입력 + 대기

        Args:
            control: pywinauto 컨트롤
            keys: 입력할 키
            sleep_after: 입력 후 대기 시간 (global_delay 적용)
            **kwargs: type_keys에 전달할 추가 인자
        """
        self.type_keys(control, keys, **kwargs)
        time.sleep(sleep_after * self.global_delay)

    def paste_text(self, control, text: str, sleep_after: float = 0.15):
        """
        클립보드 복사 후 붙여넣기 (긴 텍스트 입력 최적화)

        Args:
            control: pywinauto 컨트롤
            text: 입력할 텍스트
            sleep_after: 붙여넣기 후 대기 시간 (global_delay 적용)
        """
        # 클립보드에 텍스트 복사
        pyperclip.copy(text)
        # Ctrl+V로 붙여넣기
        self.type_keys_with_delay(control, "^v", sleep_after=sleep_after, pause=0.05)

    def copy_from_control(self, control=None):
        """
        현재 선택된 셀에서 텍스트 복사 (Ctrl+C)

        Args:
            control: 복사할 컨트롤 (기본: left_spread)

        Returns:
            복사된 텍스트
        """
        if control is None:
            control = self.left_spread
        self.type_keys_with_delay(control, "^c", pause=0.05)
        return pyperclip.paste().strip()

    def convert_nationality(self, nationality: str) -> str:
        """
        내/외국인 코드 변환

        Args:
            nationality: CSV의 내/외국인 값 (N, Y, 내, 외 등)

        Returns:
            변환된 코드 (1=내국인, 2=외국인)
        """
        val = nationality.strip().upper()

        # 내국인: N, 내, 1
        if val in ['N', '내', '내국인', '1']:
            return '1'
        # 외국인: Y, 외, 2
        elif val in ['Y', '외', '외국인', '2']:
            return '2'
        else:
            # 기본값: 내국인
            self._log("WARNING", f"알 수 없는 내/외국인 값: '{nationality}' → 내국인(1)로 처리")
            return '1'

    def clean_id_number(self, id_number: str) -> Tuple[str, str]:
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
            self._log("DEBUG", f"여권번호(3) 감지: {id_number} → {cleaned}")
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
                self._log("DEBUG", f"외국인등록번호(2) 감지: {id_number} → {cleaned}")
                return ('2', cleaned)
            else:
                self._log("DEBUG", f"주민등록번호(1): {id_number} → {cleaned}")
                return ('1', cleaned)
        else:
            self._log("WARNING", f"비정상 번호 길이 ({len(cleaned)}자리): {id_number}")
            # 기본값: 주민등록번호로 처리
            return ('1', cleaned)
