# 개발 가이드

사원등록 자동화 프로젝트 개발을 위한 가이드

## 환경 설정

### 필수 요구사항

- Python 3.8+
- Windows OS
- 사원등록 프로그램 (MFC 애플리케이션)

### 의존성 설치

```bash
# uv 사용 (권장)
uv run python main.py

# 또는 pip 사용
pip install pywinauto pywin32 pillow
```

### 프로젝트 구조

```
newgen-erp-macro/
├── docs/                    # 문서
├── test/                    # 테스트 코드
│   ├── attempt/            # 시도 스크립트
│   └── image/              # 스크린샷
├── tab_automation.py       # 탭 자동화 모듈
├── main.py                 # 메인 실행 파일
└── test.py                 # 테스트 러너
```

## 빠른 시작

### 1. 프로그램 정보 확인

```bash
uv run python main.py
```

출력 예시:
```
사원등록 윈도우 발견!
- 제목: 사원등록
- 클래스: Afx:00CD0000:0:00000000:00000000:002D0253
- HWND: 0x00080866
```

### 2. 탭 자동화 테스트

```bash
uv run python tab_automation.py
```

### 3. 커스텀 스크립트 작성

```python
from tab_automation import TabAutomation

# 1. 연결
tab_auto = TabAutomation()
tab_auto.connect()

# 2. 탭 선택
tab_auto.select_tab("부양가족정보")

# 3. 데이터 입력 (TODO)
# ...
```

## 디버깅 팁

### 1. Spy++ 사용

**윈도우 정보 확인**:
1. Spy++ 실행 (Visual Studio에 포함)
2. 사원등록 프로그램 실행
3. Spy++에서 윈도우 찾기
4. 계층 구조 및 메시지 분석

**메시지 로그 수집**:
1. Spy++ → Messages → Log Messages
2. 타겟 윈도우 선택
3. 탭 클릭
4. 로그 저장

### 2. pywinauto 디버깅

**컨트롤 트리 출력**:
```python
from pywinauto import application

app = application.Application(backend="win32")
app.connect(title="사원등록")
dlg = app.window(title="사원등록")

# 전체 컨트롤 트리
dlg.print_control_identifiers()

# 특정 컨트롤 정보
tab = dlg.child_window(class_name="Afx:TabWnd:...")
print(f"HWND: 0x{tab.handle:08X}")
print(f"위치: {tab.rectangle()}")
print(f"자식: {len(tab.children())}개")
```

**컨트롤 검색**:
```python
# 모든 하위 컨트롤
descendants = dlg.descendants()
print(f"총 {len(descendants)}개 컨트롤")

# 특정 클래스 필터링
for ctrl in descendants:
    if "TabWnd" in ctrl.class_name():
        print(f"발견: {ctrl.class_name()}")
```

### 3. 스크린샷 디버깅

**자동 캡처**:
```python
from test.capture import capture_window

hwnd = dlg.handle
capture_window(hwnd, "debug_01.png")
# → test/image/debug_01.png에 저장
```

**수동 확인**:
```python
from PIL import Image
img = Image.open("test/image/debug_01.png")
img.show()
```

### 4. 로깅

```python
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

logger.debug("탭 컨트롤 찾기 시작")
logger.info(f"HWND: 0x{hwnd:08X}")
logger.warning("탭 컨트롤을 찾을 수 없음")
logger.error("연결 실패", exc_info=True)
```

## 일반적인 문제 해결

### 문제: "사원등록 윈도우를 찾을 수 없습니다"

**원인**: 프로그램이 실행되지 않음

**해결**:
1. 사원등록 프로그램 실행 확인
2. 윈도우 제목이 정확히 "사원등록"인지 확인

### 문제: "탭 컨트롤을 찾을 수 없습니다"

**원인**: 클래스명이 변경됨

**해결**:
```python
# 부분 매칭 사용
for ctrl in dlg.descendants():
    if ctrl.class_name().startswith("Afx:TabWnd:"):
        tab_control = ctrl
        break
```

### 문제: 탭이 선택되지 않음

**원인 1**: 잘못된 좌표
```python
# ✅ 클라이언트 좌표 사용
x, y = 150, 15
```

**원인 2**: 메시지 순서
```python
# ✅ DOWN → 대기 → UP
win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, ...)
time.sleep(0.1)  # 중요!
win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, ...)
```

### 문제: 32비트 경고 메시지

```
UserWarning: 32-bit application should be automated using 32-bit Python (you use 64-bit Python)
```

**해결**: 무시 가능 (작동은 정상)

또는 32비트 Python 사용:
```bash
# 32비트 Python 다운로드
# https://www.python.org/downloads/
```

### 문제: 마우스가 움직임

**원인**: `click_input()` 사용

**해결**: `SendMessage` 사용
```python
# ❌ 마우스 움직임
tab_control.click_input()

# ✅ 마우스 고정
win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, ...)
```

## 코드 스타일

### 네이밍 규칙

```python
# 클래스: PascalCase
class TabAutomation:
    pass

# 함수/변수: snake_case
def select_tab(tab_name):
    tab_control = None

# 상수: UPPER_SNAKE_CASE
TAB_POSITIONS = {
    "기본사항": (50, 15),
}
```

### 문서화

```python
def select_tab(self, tab_name, wait_time=0.5):
    """
    특정 탭을 선택 (마우스 커서 움직이지 않음)

    Args:
        tab_name: 선택할 탭 이름 (예: "부양가족정보")
        wait_time: 탭 선택 후 대기 시간 (초)

    Returns:
        bool: 성공 여부

    Raises:
        ValueError: 알 수 없는 탭 이름
        Exception: 탭 컨트롤을 찾을 수 없음
    """
    pass
```

### 에러 처리

```python
try:
    tab_control = find_tab_control(dlg)
    select_tab(tab_control, "부양가족정보")
except Exception as e:
    logger.error(f"탭 선택 실패: {e}", exc_info=True)
    raise
```

## 다음 단계

### 1. fpUSpread80 자동화

**목표**: 부양가족 데이터 입력 자동화

**TODO**:
- [ ] Farpoint Spread 컨트롤 API 조사
- [ ] 셀 선택 방법 개발
- [ ] 데이터 입력 메서드 작성
- [ ] 여러 행 입력 처리

**참고 자료**:
- Farpoint Spread 공식 문서
- ActiveX 컨트롤 자동화 방법

### 2. 에러 처리 개선

**TODO**:
- [ ] 재시도 로직 추가
- [ ] 상세한 에러 메시지
- [ ] 로깅 강화
- [ ] 타임아웃 처리

### 3. 사용자 인터페이스

**TODO**:
- [ ] CLI 인터페이스 개선
- [ ] 설정 파일 지원 (YAML/JSON)
- [ ] 진행 상황 표시
- [ ] 결과 보고서 생성

### 4. 테스트 확장

**TODO**:
- [ ] 단위 테스트 작성
- [ ] 통합 테스트
- [ ] 에러 시나리오 테스트
- [ ] CI/CD 설정

## 참고 자료

### 문서
- [pywinauto 공식 문서](https://pywinauto.readthedocs.io/)
- [Win32 API 레퍼런스](https://docs.microsoft.com/en-us/windows/win32/api/)
- [Farpoint Spread 문서](https://www.grapecity.com/spread)

### 도구
- **Spy++**: 윈도우 메시지 분석
- **AccChecker**: 접근성 트리 확인
- **pywinauto Inspector**: pywinauto 전용 도구

### 프로젝트 문서
- [overview.md](overview.md) - 프로젝트 개요
- [window-architecture.md](window-architecture.md) - 윈도우 구조
- [tab-automation.md](tab-automation.md) - 탭 자동화
- [testing-framework.md](testing-framework.md) - 테스트 프레임워크

## 기여 가이드

### 새 Attempt 추가

1. `test/attempt/attemptN_description.py` 생성
2. `run(dlg, capture_func)` 함수 구현
3. `test.py`에 추가
4. 스크린샷 촬영 및 결과 확인
5. 문서 업데이트

### 버그 리포트

다음 정보 포함:
- Python 버전
- OS 버전
- 사원등록 프로그램 버전
- 에러 메시지 전문
- 재현 단계

### 코드 리뷰 체크리스트

- [ ] 코드 스타일 준수
- [ ] 문서화 완료
- [ ] 에러 처리 추가
- [ ] 테스트 작성
- [ ] 스크린샷 첨부 (UI 변경 시)
