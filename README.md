# 부양가족 대량 입력 자동화

사원등록 프로그램에서 부양가족 정보를 CSV 파일로부터 자동으로 대량 입력하는 GUI 애플리케이션입니다.

**마우스 커서를 움직이지 않고** 데이터 입력을 자동화하며, **Pause 키로 안전하게 중지**할 수 있습니다.

## 🎯 프로젝트 개요

이 프로젝트는 **사원등록** MFC 애플리케이션의 UI 요소를 자동으로 제어하여:
- ✅ 탭 선택 (마우스 움직임 없음) - **완료**
- ✅ 직원 정보 입력 (사번, 성명, 주민번호) - **완료**
- ✅ 부양가족 데이터 대량 입력 - **완료**
- ✅ GUI 애플리케이션 - **완료**
- ✅ 실행파일 빌드 - **완료**

## 주요 기능

- **GUI 인터페이스**: CustomTkinter 기반의 사용자 친화적 인터페이스
- **CSV 파일 기반**: 엑셀에서 편집 가능한 CSV 형식 지원
- **Pause 키 중지**: Pause 키 3번으로 안전하게 중지 (마우스 불필요)
- **실시간 로그**: 처리 상황 실시간 모니터링
- **Dry Run 모드**: 실제 입력 없이 테스트 가능
- **부분 실행**: 처리할 사원 수 지정 가능
- **플로팅 안내 창**: 실행 중 중지 방법 안내
- **실행파일 제공**: Python 설치 없이 .exe 파일로 실행 가능

## 시작 전 준비사항

### 프로그램 초기 상태

자동화 스크립트를 실행하기 전에 **사원등록** 프로그램이 다음 상태로 시작되어야 합니다:

![초기 화면](docs/images/initial-screen.png)

**필수 조건:**
- ✅ "부양가족정보" 탭이 선택된 상태
- ✅ 왼쪽 사원 목록이 비어있거나 첫 번째 행이 선택된 상태
- ✅ 오른쪽 부양가족 입력 영역이 비어있는 상태

## 빠른 시작

### ⭐ 실행파일 사용 (권장)

1. **실행파일 빌드** (최초 1회만)
   ```bash
   build.bat
   ```

2. **실행파일 실행**
   - `dist/부양가족_대량입력.exe` 더블 클릭
   - CSV 파일 선택
   - 옵션 설정 후 **▶ 시작** 버튼 클릭

3. **중지 방법**
   - **Pause 키를 3번 연속으로 누르세요** (2초 이내)
   - 플로팅 안내 창에 중지 방법 표시됨

### 📝 CSV 파일 형식

| 컬럼명 | 설명 | 예시 |
|--------|------|------|
| 사번 | 사원 번호 | 20240001 |
| 성명 | 사원 이름 | 홍길동 |
| 부양가족성명 | 부양가족 이름 | 홍아들 |
| 관계코드 | 연말정산 관계 코드 | 4 (자녀) |
| 내외국인 | N=내국인, Y=외국인 | N |
| 주민번호 | 주민등록번호 | 001231-1234567 |
| 기본공제 | Y=공제, N=공제안함 | Y |
| 자녀공제 | Y=공제, N=공제안함 | Y |

### Python으로 직접 실행

```bash
# 의존성 설치
uv sync

# GUI 실행
uv run python gui_app.py

# CLI 실행 (고급 사용자)
uv run python bulk_dependent_input.py --csv "데이터.csv" --count 5
```

### 탭 자동화 (기존 좌표 기반)

```python
from src.tab_automation import TabAutomation

# 1. 연결
tab_auto = TabAutomation()
tab_auto.connect()

# 2. 탭 선택
tab_auto.select_tab("부양가족정보")
```

### 직원 정보 입력

```python
from src.employee_input import EmployeeInput
from src.tab_automation import TabAutomation

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
# 예제 실행 (성공한 방법들)
uv run python examples/example_employee_selection.py
uv run python examples/example_tab_switching.py
uv run python examples/example_data_input.py

# 전체 테스트
uv run python test.py
```

## 프로젝트 구조

```
newgen-erp-macro/
├── gui_app.py                     # 🎯 GUI 메인 애플리케이션
├── bulk_dependent_input.py        # 부양가족 대량 입력 자동화 로직
├── gui_app.spec                   # PyInstaller 빌드 설정
├── build.bat                      # 실행파일 빌드 스크립트
├── dist/                          # 빌드 결과물
│   └── 부양가족_대량입력.exe      # ⭐ 실행파일
├── src/                           # 핵심 자동화 모듈
│   ├── tab_automation.py          # 탭 자동화 모듈
│   ├── employee_input.py          # 직원 정보 입력 모듈
│   ├── csv_reader.py              # CSV 파일 리더
│   └── ...                        # 기타 유틸리티
├── examples/                      # 성공한 자동화 예제
│   ├── example_employee_selection.py
│   ├── example_tab_switching.py
│   └── example_data_input.py
├── docs/                          # 📚 상세 문서
│   ├── automation-guide.md        # 완전한 자동화 가이드
│   └── ...                        # 기타 문서
├── test/                          # 🧪 테스트 및 개발 이력
│   ├── attempt/                   # 53개 시도 스크립트들
│   └── ...
├── logs/                          # 실행 로그 저장소
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
from src.tab_automation import TabAutomation

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
# ✅ 성공한 자동화 예제 실행
uv run python examples/example_employee_selection.py
uv run python examples/example_tab_switching.py
uv run python examples/example_data_input.py

# 전체 테스트 스위트
uv run python test.py

# Spy++와 함께 실시간 모니터링 테스트
uv run python test_with_spy.py

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

### 기본 모니터링 (`src/message_monitor.py`)

```python
from src.message_monitor import MessageMonitor

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

### 고급 모니터링 (`src/advanced_message_monitor.py`)

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
from src.employee_input import EmployeeInput
from src.tab_automation import TabAutomation

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

### 실행 의존성
- **customtkinter**: GUI 프레임워크
- **keyboard**: 키보드 이벤트 처리 (Pause 키 감지)
- **pywinauto**: Windows UI 자동화
- **pyperclip**: 클립보드 작업
- **pywin32**: Windows API

### 개발 의존성
- **pyinstaller**: 실행파일 빌드

전체 의존성은 `pyproject.toml` 참조

## 트러블슈팅

### keyboard 패키지 관련 오류
- **증상**: 프로그램이 실행되지 않거나 Pause 키가 작동하지 않음
- **해결**: 관리자 권한으로 실행
  - 실행파일: 우클릭 → "관리자 권한으로 실행"
  - Python: PowerShell을 관리자 권한으로 실행 후 `uv run python gui_app.py`

### 사원 정보를 찾을 수 없음
- **증상**: "CSV 데이터 없음, 건너뜀"
- **해결**: CSV의 사번과 화면의 사번이 일치하는지 확인

### 입력이 느림
- 정상입니다. 안정성을 위해 각 입력 후 대기 시간이 있습니다
- 평균 1명당 5-10초 소요

### 32비트/64비트 경고

```
UserWarning: 32-bit application should be automated using 32-bit Python
```

이 경고는 무시해도 됩니다. 64비트 Python으로도 32비트 애플리케이션 제어가 가능합니다.

## 📚 문서

### 🎯 시작하기
- **[자동화 가이드](docs/automation-guide.md)** ⭐ - **완전한 좌표 독립적 자동화 가이드** (53회 시도 결과)
  - 사원 선택 (좌표 없음)
  - 탭 전환 (좌표 없음)
  - 데이터 입력 (좌표 없음)
  - 완전한 자동화 예제
  - 트러블슈팅

- **[성공한 방법](docs/successful-method.md)** - 최종 성공 방법 정리 (Attempt 43, 52, 53)

### 📖 상세 문서

#### 기본 개념
- **[프로젝트 개요](docs/overview.md)** - 목표, 제약사항, 현재 상태
- **[윈도우 구조](docs/window-architecture.md)** - Spy++ 분석, 컨트롤 정보

#### 자동화 모듈
- **[탭 자동화](docs/tab-automation.md)** - 탭 선택 방법 (기존 좌표 기반)
- **[사원 입력](docs/employee-input.md)** - 직원 정보 입력 가이드

#### 개발 및 테스트
- **[테스트 프레임워크](docs/testing-framework.md)** - attempt 패턴, 스크린샷 검증
- **[개발 가이드](docs/development-guide.md)** - 디버깅, 트러블슈팅
- **[Spy++ 실시간 모니터링](docs/spy-realtime-monitoring.md)** - 메시지 모니터링 도구

### 🔍 빠른 참조
- **좌표 없는 사원 선택**: [automation-guide.md](docs/automation-guide.md#사원-선택-좌표-없음)
- **좌표 없는 탭 전환**: [automation-guide.md](docs/automation-guide.md#탭-전환-좌표-없음)
- **데이터 입력**: [automation-guide.md](docs/automation-guide.md#데이터-입력-좌표-없음)
- **완전한 자동화 예제**: [automation-guide.md](docs/automation-guide.md#완전한-자동화-예제)
- **트러블슈팅**: [automation-guide.md](docs/automation-guide.md#트러블슈팅)

## 라이선스

이 프로젝트는 내부 사용 목적으로 개발되었습니다.

## 작성자

손기령 (giryeong@kodebox.io)
