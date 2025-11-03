"""
스프레드 컨트롤러 모듈

pywinauto를 이용한 스프레드 찾기 및 제어 기능을 제공합니다.
"""

from typing import Optional


class SpreadController:
    """스프레드 제어 클래스"""

    def __init__(self, dialog, log_callback=None):
        """
        초기화

        Args:
            dialog: pywinauto dialog wrapper
            log_callback: 로그 출력 콜백 함수 (level, message)
        """
        self.dialog = dialog
        self.log_callback = log_callback
        self.left_spread = None
        self.right_spread = None

    def _log(self, level: str, message: str):
        """로그 출력 (콜백이 있으면 사용)"""
        if self.log_callback:
            self.log_callback(level, message)

    def find_spreads(self) -> tuple:
        """
        왼쪽(사원 목록)과 오른쪽(부양가족 상세) 스프레드 찾기

        Returns:
            (left_spread, right_spread) 튜플

        Raises:
            Exception: 스프레드를 찾을 수 없는 경우
        """
        spreads = self.dialog.children(class_name="fpUSpread80")

        if not spreads:
            error_msg = "❌ 스프레드 컨트롤을 찾을 수 없습니다!\n\n"
            error_msg += "다음을 확인하세요:\n"
            error_msg += "1. 사원등록 창이 제대로 열려있는지\n"
            error_msg += "2. '부양가족정보' 탭이 선택되어 있는지"
            raise Exception(error_msg)

        if len(spreads) < 2:
            error_msg = "❌ 부양가족상세 탭이 열려있지 않습니다!\n\n"
            error_msg += "다음을 확인하세요:\n"
            error_msg += "1. 사원등록 창에서 '부양가족정보' 탭을 선택했는지\n"
            error_msg += "2. 오른쪽 부양가족상세 영역이 표시되어 있는지\n\n"
            error_msg += f"(현재 스프레드 개수: {len(spreads)}개, 필요: 2개 이상)"
            raise Exception(error_msg)

        # X 좌표로 정렬
        spreads.sort(key=lambda s: s.rectangle().left)

        self.left_spread = spreads[0]
        self.right_spread = spreads[-1]

        return self.left_spread, self.right_spread

    def verify_spreads(self):
        """
        스프레드 상태 재확인

        Raises:
            Exception: 스프레드가 유효하지 않은 경우
        """
        spreads = self.dialog.children(class_name="fpUSpread80")

        if len(spreads) < 2:
            error_msg = "❌ 부양가족상세 탭이 열려있지 않습니다!\n\n"
            error_msg += "다음을 확인하세요:\n"
            error_msg += "1. 사원등록 창에서 '부양가족정보' 탭을 선택했는지\n"
            error_msg += "2. 오른쪽 부양가족상세 영역이 표시되어 있는지\n\n"
            error_msg += f"(현재 스프레드 개수: {len(spreads)}개, 필요: 2개 이상)"
            raise Exception(error_msg)

    def get_left_spread(self):
        """왼쪽 스프레드 반환"""
        return self.left_spread

    def get_right_spread(self):
        """오른쪽 스프레드 반환"""
        return self.right_spread
