# 사원등록 자동화

pywinauto를 사용한 Windows 사원등록 프로그램 자동화 도구

**마우스 커서를 움직이지 않고** 탭 선택 및 데이터 입력을 자동화합니다.

## 프로젝트 개요

이 프로젝트는 **사원등록** MFC 애플리케이션의 UI 요소를 자동으로 제어하여:
- ✅ 탭 선택 (마우스 움직임 없음) - **완료**
- ✅ 직원 정보 입력 (사번, 성명, 주민번호) - **완료**
- 🔄 부양가족 데이터 입력 - **개발 중**

## 주요 기능

- **탭 자동화**: Win32 SendMessage를 사용한 안정적인 탭 선택
- **마우스 비침범**: 물리적 마우스 이동 없이 윈도우 메시지만 사용
- **스크린샷 검증**: 모든 작업 단계를 이미지로 캡처하여 검증
- **체계적인 테스트**: attempt 패턴으로 다양한 방법 시도 및 문서화

## 빠른 시작

### 탭 자동화

```python
from tab_automation import TabAutomation

# 1. 연결
tab_auto = TabAutomation()
tab_auto.connect()

# 2. 탭 선택
tab_auto.select_tab("부양가족정보")
```

### 직원 정보 입력

```python
from employee_input import EmployeeInput
from tab_automation import TabAutomation

# 1. 연결
emp_input = EmployeeInput()
emp_input.connect()

# 2. 기본사항 탭 선택
tab_auto = TabAutomation()
tab_auto.connect()
tab_auto.select_tab("기본사항")

# 3. 직원 정보 입력
result = emp_input.input_employee(
    employee_no="2025100",
    name="테스트사원",
    id_number="920315-1234567",
    age="33"
)

print(f"성공: {result['success_count']}/{result['total']}개")
```

### 테스트 실행

```bash
# 탭 자동화 테스트
uv run python tab_automation.py

# 전체 테스트
uv run python test.py
```

## 프로젝트 구조

```
newgen-erp-macro/
├── docs/                          # 📚 문서
│   ├── overview.md                # 프로젝트 개요
│   ├── window-architecture.md     # 윈도우 구조 분석
│   ├── tab-automation.md          # 탭 자동화 가이드
│   ├── testing-framework.md       # 테스트 프레임워크
│   ├── development-guide.md       # 개발 가이드
│   └── spy-realtime-monitoring.md # Spy++ 실시간 모니터링
├── test/                          # 🧪 테스트 모듈
│   ├── attempt/                   # 시도 스크립트들
│   │   ├── attempt01_click_children.py
│   │   ├── attempt06_direct_tab_hwnd.py  # ✅ 성공
│   │   └── attempt07_robust_tab_find.py
│   ├── image/                     # 스크린샷 저장소
│   ├── message_log_*.txt          # 메시지 모니터링 로그
│   └── capture.py                 # 캡처 유틸리티
├── tab_automation.py              # 🎯 탭 자동화 모듈
├── employee_input.py              # 👤 직원 정보 입력 모듈
├── message_monitor.py             # 🔍 기본 메시지 모니터링
├── advanced_message_monitor.py    # 🔬 고급 메시지 모니터링
├── test_with_spy.py               # Spy++ 연동 테스트
├── test_employee_input_with_monitoring.py  # 직원 입력 + 모니터링
├── analyze_basic_tab.py           # 기본사항 탭 분석 도구
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

### 탭 자동화 모듈 사용

```python
from tab_automation import TabAutomation

# 객체 생성 및 연결
tab_auto = TabAutomation()
tab_auto.connect()

# 탭 선택 (이름으로)
tab_auto.select_tab("부양가족정보")
tab_auto.select_tab("소득자료")

# 탭 선택 (인덱스로)
tab_auto.select_tab_by_index(1)  # 0=기본사항, 1=부양가족정보, 2=소득자료
```

### 테스트 실행

```bash
# 탭 자동화 데모
uv run python tab_automation.py

# 전체 테스트 스위트
uv run python test.py

# Spy++와 함께 실시간 모니터링 테스트
uv run python test_with_spy.py

# 기본 메시지 모니터링 (SendMessage만)
uv run python message_monitor.py

# 고급 메시지 모니터링 (멀티스레드 + 로그 파일)
uv run python advanced_message_monitor.py

# 직원 정보 입력
uv run python employee_input.py

# 직원 정보 입력 + 메시지 모니터링
uv run python test_employee_input_with_monitoring.py

# 프로그램 정보 확인
uv run python main.py
```

### 스크린샷 확인

테스트 실행 후 `test/image/` 폴더에서 결과 스크린샷을 확인할 수 있습니다.

### 메시지 로그 확인

메시지 모니터링 실행 후 `test/message_log_*.txt`에서 상세 로그를 확인할 수 있습니다:
- 타임스탬프 (밀리초 단위)
- HWND, 메시지 코드, 파라미터
- 좌표 (x, y) 자동 디코딩

## 메시지 모니터링

### 기본 모니터링 (`message_monitor.py`)

```python
from message_monitor import MessageMonitor

# 모니터 생성
monitor = MessageMonitor(target_hwnd=tab_hwnd)
monitor.start()

# 자동화 실행
# ... SendMessage 호출 ...

monitor.stop()

# 결과 확인
messages = monitor.get_messages()
```

**출력 예시:**
```
📤 SEND: [12:04:44.020] HWND=0x000608EE WM_LBUTTONDOWN wParam=0x00000001 lParam=0x000F0096 (x=150, y=15)
📤 SEND: [12:04:44.256] HWND=0x000608EE WM_LBUTTONUP wParam=0x00000000 lParam=0x000F0096 (x=150, y=15)
```

### 고급 모니터링 (`advanced_message_monitor.py`)

- ✅ 멀티스레딩: 모니터링과 자동화 동시 실행
- ✅ 로그 파일 자동 저장 (`test/message_log_*.txt`)
- ✅ 밀리초 단위 타임스탬프
- ✅ 메시지 분석 및 통계

**특징:**
- SendMessage 호출을 직접 로깅 (100% 정확)
- 메시지 필터링 (관심있는 것만)
- LPARAM 좌표 자동 디코딩
- 스레드 안전 로깅

**한계:**
- 시스템 내부 메시지 (WM_NOTIFY 등)는 캡처 안 됨
- 완전한 후킹은 DLL 인젝션 필요
- 하지만 디버깅에는 충분함

## 직원 정보 입력

### 기본 사용법

```python
from employee_input import EmployeeInput
from tab_automation import TabAutomation

# 1. 연결
emp_input = EmployeeInput()
emp_input.connect()

# 2. 기본사항 탭 선택
tab_auto = TabAutomation()
tab_auto.connect()
tab_auto.select_tab("기본사항")

# 3. Spread 컨트롤 찾기
emp_input.find_spread_control()

# 4. 직원 정보 입력
result = emp_input.input_employee(
    employee_no="2025100",
    name="테스트사원",
    id_number="920315-1234567",
    age="33"
)

if result['success']:
    print(f"✅ {result['success_count']}/{result['total']}개 입력 완료")
```

### 확립된 방법 (Attempt 18)

**핵심 원리:**
1. fpUSpread80 Spread #2 (왼쪽 직원 목록) 찾기
2. WM_LBUTTONDOWN/UP로 셀 클릭하여 선택
3. WM_CHAR로 각 문자 입력
4. VK_RETURN으로 Enter 키 전송

**입력 가능 필드 및 좌표 매핑:**
- 사번 (employee_no): x=50, y=30
- 성명 (name): x=100, y=30
- 주민번호 (id_number): x=200, y=30
- 나이 (age): x=320, y=30

**메시지 시퀀스:**
```
각 필드마다:
1. WM_LBUTTONDOWN/UP (x, y) → 셀 선택
2. WM_CHAR (각 문자) → 텍스트 입력
3. WM_KEYDOWN/UP (VK_RETURN) → Enter 키
```

**왜 fpUSpread80인가?**
- 기본사항 탭의 왼쪽 직원 목록은 fpUSpread80 ActiveX 스프레드시트 컨트롤
- Edit 컨트롤(SPR32DU80)은 숨겨져 있고 실제로는 작동하지 않음
- Spread #2 (세 번째 fpUSpread80)가 왼쪽 직원 목록

### 실패한 방법들

- ❌ SPR32DU80EditHScroll 컨트롤 사용 - 숨겨진 컨트롤, 실제 입력 안 됨
- ❌ set_edit_text() 직접 사용 - 권한 오류
- ❌ SetFocus() - 액세스 거부
- ❌ WM_SETTEXT만 사용 - 변경 감지 안 됨

### 발견 과정

18회의 시도를 통해 올바른 방법 확립:
1. Attempt 08-09: SPR32DU80 시도 → 사용자 확인 결과 실패
2. Attempt 11: WindowFromPoint로 fpUSpread80 발견
3. Attempt 12-13: fpUSpread80 입력 성공 확인
4. Attempt 15: 3개의 Spread 컨트롤 발견
5. Attempt 17: 좌표 매핑 완료
6. Attempt 18: 완전한 직원 정보 입력 성공

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

- ❌ 물리적 마우스 이동 금지 (pyautogui, mouse 등)
- ✅ Win32 SendMessage만 사용
- ✅ 마우스 커서 움직이지 않음

자세한 내용은 [docs/tab-automation.md](docs/tab-automation.md) 참조

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

## 📚 문서

### 주요 문서
- **[프로젝트 개요](docs/overview.md)** - 목표, 제약사항, 현재 상태
- **[탭 자동화](docs/tab-automation.md)** - 성공한 방법, 실패 사례, 코드 예제
- **[윈도우 구조](docs/window-architecture.md)** - Spy++ 분석, 컨트롤 정보
- **[테스트 프레임워크](docs/testing-framework.md)** - attempt 패턴, 스크린샷 검증
- **[개발 가이드](docs/development-guide.md)** - 디버깅, 트러블슈팅, 다음 단계
- **[Spy++ 실시간 모니터링](docs/spy-realtime-monitoring.md)** - 자동화 스크립트와 Spy++ 동시 실행

### 빠른 참조
- 탭 선택 방법: [tab-automation.md](docs/tab-automation.md#성공한-방법-2025-10-30)
- 컨트롤 찾기: [window-architecture.md](docs/window-architecture.md#안정적인-컨트롤-찾기)
- 디버깅 팁: [development-guide.md](docs/development-guide.md#디버깅-팁)

## 라이선스

이 프로젝트는 내부 사용 목적으로 개발되었습니다.

## 작성자

손기령 (giryeong@kodebox.io)
