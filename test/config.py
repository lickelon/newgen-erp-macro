"""
테스트 설정 로더

민감정보를 하드코딩하지 않고 test_config.json에서 읽어옵니다.
"""
import json
import os
from pathlib import Path


class TestConfig:
    """테스트 설정을 로드하고 관리하는 클래스"""

    def __init__(self, config_path=None):
        """
        Args:
            config_path: 설정 파일 경로 (None이면 프로젝트 루트의 test_config.json)
        """
        if config_path is None:
            # 프로젝트 루트 찾기
            current = Path(__file__).parent.parent
            config_path = current / "test_config.json"

        self.config_path = Path(config_path)
        self._data = None
        self._load()

    def _load(self):
        """설정 파일 로드"""
        if not self.config_path.exists():
            raise FileNotFoundError(
                f"테스트 설정 파일이 없습니다: {self.config_path}\n"
                f"test_config.example.json을 test_config.json으로 복사하고 실제 값을 입력하세요."
            )

        with open(self.config_path, 'r', encoding='utf-8') as f:
            self._data = json.load(f)

    def get(self, key_path, default=None):
        """
        점(.) 구분자로 중첩된 키에 접근

        Args:
            key_path: "test_data.employee.empno" 형식의 키 경로
            default: 키가 없을 때 반환할 기본값

        Returns:
            설정값

        Examples:
            >>> config = TestConfig()
            >>> empno = config.get("test_data.employee.empno")
            >>> resident = config.get("test_data.resident_numbers.with_hyphen")
        """
        keys = key_path.split('.')
        value = self._data

        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default

        return value

    @property
    def empno(self):
        """테스트용 사번"""
        return self.get("test_data.employee.empno")

    @property
    def employee_name(self):
        """테스트용 사원 이름"""
        return self.get("test_data.employee.name")

    @property
    def resident_number_with_hyphen(self):
        """하이픈 포함 주민등록번호"""
        return self.get("test_data.resident_numbers.with_hyphen")

    @property
    def resident_number_without_hyphen(self):
        """하이픈 없는 주민등록번호"""
        return self.get("test_data.resident_numbers.without_hyphen")


# 싱글톤 인스턴스
_config_instance = None


def get_test_config():
    """
    테스트 설정 싱글톤 인스턴스 반환

    Returns:
        TestConfig 인스턴스

    Examples:
        >>> from test.config import get_test_config
        >>> config = get_test_config()
        >>> empno = config.empno
    """
    global _config_instance

    if _config_instance is None:
        _config_instance = TestConfig()

    return _config_instance


# 편의 함수들
def get_empno():
    """테스트용 사번 반환"""
    return get_test_config().empno


def get_resident_number(with_hyphen=True):
    """
    테스트용 주민등록번호 반환

    Args:
        with_hyphen: True이면 하이픈 포함, False면 하이픈 없음

    Returns:
        주민등록번호 문자열
    """
    config = get_test_config()
    if with_hyphen:
        return config.resident_number_with_hyphen
    else:
        return config.resident_number_without_hyphen


if __name__ == "__main__":
    # 테스트
    try:
        config = get_test_config()
        print("설정 로드 성공!")
        print(f"사번: {config.empno}")
        print(f"이름: {config.employee_name}")
        print(f"주민등록번호 (하이픈 포함): {config.resident_number_with_hyphen}")
        print(f"주민등록번호 (하이픈 없음): {config.resident_number_without_hyphen}")
    except FileNotFoundError as e:
        print(f"❌ {e}")
