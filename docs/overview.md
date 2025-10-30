# 사원등록 자동화 - 프로젝트 개요

## 목표

**사원등록** 프로그램에서 **탭**을 자동으로 선택하고 **부양가족 데이터**를 입력

## 제약사항

### 1. 마우스 직접 이동 금지
- ❌ `pyautogui.click()` - 물리적 마우스 이동
- ❌ `mouse.click()` - 물리적 마우스 이동
- ❌ 절대 좌표 물리적 클릭
- ✅ `pywinauto`의 `click_input()` - 윈도우 메시지 사용
- ✅ `win32api.SendMessage()` - 윈도우 메시지 직접 전송

**이유**: 물리적 마우스 이동은 사용자의 작업을 방해하고 신뢰성이 낮음

### 2. pywinauto 사용
- MFC 기반 애플리케이션이므로 win32 백엔드 사용
- 32비트 애플리케이션이지만 64비트 Python으로 제어 가능 (경고 무시)

### 3. 윈도우 메시지 방식
- 마우스 커서를 움직이지 않고 윈도우 핸들에 직접 메시지 전송
- 백그라운드에서 작동 가능

## 테스트 방법

### 1. 스크린샷 기반 평가
- 매 단위 실행마다 **스크린샷 촬영**
- Claude가 **직접 이미지 확인**하여 탭 선택 여부 평가
- 성공할 때까지 **계속 시도**

### 2. 반복적 시도
- 다양한 접근 방법을 순차적으로 시도
- 각 시도는 독립적인 스크립트로 관리
- 실패 시 다음 방법으로 진행

## 현재 상태 (2025-10-30)

### ✅ 완료
1. ✅ 사원등록 윈도우 연결
2. ✅ 탭 선택 자동화 (마우스 움직임 없이)
3. ✅ Spy++ 분석으로 윈도우 계층 구조 파악
4. ✅ fpUSpread80 부양가족 입력 표 발견
5. ✅ 안정적인 탭 컨트롤 찾기 (부분 매칭)
6. ✅ **직원 정보 입력 자동화 완료** (18회 시도 끝에 성공)
   - fpUSpread80 Spread #2 (왼쪽 직원 목록) 사용
   - 4개 필드 좌표 매핑 완료 (사번, 성명, 주민번호, 나이)
   - 프로덕션 모듈 `employee_input.py` 작성
   - 상세 문서 작성 ([employee-input.md](employee-input.md))

### 🔄 진행 중
1. **부양가족 데이터 입력 자동화**
   - 부양가족정보 탭의 fpUSpread80 컨트롤 조사
   - 유사한 방법 적용 예정

## 문서 구조

- **[overview.md](overview.md)** (이 문서) - 프로젝트 개요
- **[window-architecture.md](window-architecture.md)** - 윈도우 구조 분석
- **[tab-automation.md](tab-automation.md)** - 탭 자동화 방법
- **[employee-input.md](employee-input.md)** - 직원 정보 입력 자동화 ✨새로운 문서
- **[testing-framework.md](testing-framework.md)** - 테스트 프레임워크
- **[development-guide.md](development-guide.md)** - 개발 가이드

## 빠른 시작

### 탭 선택

```python
from tab_automation import TabAutomation

# 탭 자동화 객체 생성
tab_auto = TabAutomation()

# 연결
tab_auto.connect()

# 부양가족정보 탭 선택
tab_auto.select_tab("부양가족정보")
```

자세한 내용은 [tab-automation.md](tab-automation.md)를 참고하세요.

### 직원 정보 입력

```python
from employee_input import EmployeeInput
from tab_automation import TabAutomation

# 연결 및 탭 선택
emp_input = EmployeeInput()
emp_input.connect()

tab_auto = TabAutomation()
tab_auto.connect()
tab_auto.select_tab("기본사항")

# 직원 정보 입력
result = emp_input.input_employee(
    employee_no="2025100",
    name="테스트사원",
    id_number="920315-1234567",
    age="33"
)

print(f"✅ {result['success_count']}/4개 입력 완료")
```

자세한 내용은 [employee-input.md](employee-input.md)를 참고하세요.
