# 연말정산 추가자료입력 자동화

pywinauto를 사용한 Windows 연말정산 프로그램 자동화 도구

## 프로젝트 개요

이 프로젝트는 **연말정산추가자료입력** MFC 애플리케이션의 UI 요소를 자동으로 제어하여 부양가족 탭 선택 등의 작업을 자동화합니다.

## 주요 기능

- Windows UI 자동화 (pywinauto 기반)
- 스크린샷 기반 테스트 검증
- 체계적인 시도 방법론 (attempt 패턴)

## 프로젝트 구조

```
newgen-erp-macro/
├── docs/                          # 문서
│   └── testing-strategy.md        # 테스트 전략 문서
├── test/                          # 테스트 모듈
│   ├── attempt/                   # 시도 스크립트들
│   │   ├── attempt01_click_children.py
│   │   └── ...
│   ├── image/                     # 스크린샷 저장소
│   └── capture.py                 # 캡처 유틸리티
├── archive/                       # 구버전 파일 보관
├── main.py                        # 메인 자동화 스크립트
├── test.py                        # 테스트 실행 스크립트
├── pyproject.toml                 # 프로젝트 설정 (uv)
└── README.md                      # 이 파일
```

## 설치

### 요구사항

- Python 3.10+
- Windows OS
- uv (Python 패키지 관리자)

### uv 설치

```bash
# Windows (PowerShell)
irm https://astral.sh/uv/install.ps1 | iex
```

### 프로젝트 설정

```bash
# 저장소 클론
cd newgen-erp-macro

# 의존성 설치
uv sync
```

## 사용법

### 기본 테스트 실행

```bash
# test.py 실행 (현재 활성화된 attempt 실행)
uv run python test.py
```

### 메인 자동화 실행

```bash
# main.py 실행 (최종 자동화 스크립트)
uv run python main.py
```

### 스크린샷 확인

테스트 실행 후 `test/image/` 폴더에서 결과 스크린샷을 확인할 수 있습니다.

## 개발 가이드

### 새로운 Attempt 추가

1. `test/attempt/` 폴더에 새 파일 생성
   ```python
   # test/attempt/attemptXX_description.py

   def run(dlg, capture_func):
       """
       Args:
           dlg: pywinauto 윈도우 객체
           capture_func: 스크린샷 함수 (filename) -> None

       Returns:
           dict: {"success": bool, "message": str}
       """
       # 구현
       pass
   ```

2. `test.py`에서 import하여 실행
   ```python
   from test.attempt.attemptXX_description import run as attemptXX
   result = attemptXX(dlg, capture_func)
   ```

### 제약사항

- ❌ 물리적 마우스 이동 금지 (pyautogui 등)
- ✅ pywinauto의 `click_input()` 사용
- ✅ 윈도우 메시지 방식만 사용

자세한 내용은 [docs/testing-strategy.md](docs/testing-strategy.md) 참조

## 의존성

- **pywinauto**: Windows UI 자동화
- **mss**: 스크린샷 캡처
- **Pillow**: 이미지 처리
- **pywin32**: Windows API

전체 의존성은 `pyproject.toml` 참조

## 트러블슈팅

### 32비트/64비트 경고

```
UserWarning: 32-bit application should be automated using 32-bit Python
```

이 경고는 무시해도 됩니다. 64비트 Python으로도 32비트 애플리케이션 제어가 가능합니다.

### 권한 오류

일부 작업은 관리자 권한이 필요할 수 있습니다. PowerShell 또는 터미널을 관리자 권한으로 실행하세요.

## 문서

- [테스트 전략](docs/testing-strategy.md) - 전체 테스트 전략 및 방법론
- [Attempt 패턴](docs/testing-strategy.md#attempt-스크립트-구조) - 시도 스크립트 작성 가이드

## 라이선스

이 프로젝트는 내부 사용 목적으로 개발되었습니다.

## 작성자

손기령 (giryeong@kodebox.io)
