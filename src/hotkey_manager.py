"""
핫키 관리 모듈

Pause 키를 이용한 중지 기능을 제공합니다.
"""

import time
import keyboard


class HotkeyManager:
    """핫키 관리 클래스"""

    def __init__(self, log_callback=None):
        """
        초기화

        Args:
            log_callback: 로그 출력 콜백 함수 (level, message)
        """
        self.log_callback = log_callback
        self.stop_requested = False
        self.pause_press_count = 0
        self.last_pause_time = 0

        # Pause 키 리스너 등록
        keyboard.on_press_key('pause', self._on_pause_press)

    def _log(self, level: str, message: str):
        """로그 출력 (콜백이 있으면 사용)"""
        if self.log_callback:
            self.log_callback(level, message)

    def _on_pause_press(self, event):
        """
        Pause 키 눌림 콜백

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
            self._log("WARNING", "Pause 키 3번 감지 - 중지 요청")
            self.stop_requested = True
            self.pause_press_count = 0  # 리셋
        else:
            self._log("INFO", f"Pause 키 감지 ({self.pause_press_count}/3)")

    def is_stop_requested(self) -> bool:
        """
        중지 요청 상태 확인

        Returns:
            True면 중지 요청됨, False면 계속 진행
        """
        return self.stop_requested

    def request_stop(self):
        """중지 요청"""
        self._log("WARNING", "중지 요청됨 - 현재 작업 완료 후 중단...")
        self.stop_requested = True

    def cleanup(self):
        """리소스 정리 (keyboard 후크 해제)"""
        try:
            keyboard.unhook_key('pause')
            self._log("INFO", "Pause 키 리스너 해제됨")
        except Exception as e:
            self._log("WARNING", f"Pause 키 리스너 해제 실패: {e}")
