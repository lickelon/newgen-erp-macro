# 직원 정보 입력 자동화

사원등록 프로그램의 기본사항 탭에서 직원 정보를 자동으로 입력하는 방법을 설명합니다.

## 목차

1. [개요](#개요)
2. [확립된 방법](#확립된-방법)
3. [좌표 매핑](#좌표-매핑)
4. [코드 예제](#코드-예제)
5. [발견 과정](#발견-과정)
6. [실패한 방법들](#실패한-방법들)
7. [트러블슈팅](#트러블슈팅)

## 개요

**목표**: 사원등록 프로그램의 왼쪽 직원 목록에 새 직원 정보 추가

**입력 필드**:
- 사번 (Employee Number)
- 성명 (Name)
- 주민번호 (ID Number)
- 나이 (Age)

**제약사항**:
- ✅ 마우스 커서를 움직이지 않음
- ✅ Win32 SendMessage만 사용
- ✅ 모든 작업을 윈도우 메시지로 처리

## 확립된 방법

### 핵심 원리 (Attempt 18)

왼쪽 직원 목록은 **fpUSpread80** ActiveX 스프레드시트 컨트롤입니다.

**단계**:
1. fpUSpread80 Spread #2 (왼쪽 목록) 찾기
2. 각 필드마다:
   - WM_LBUTTONDOWN/UP로 셀 클릭
   - WM_CHAR로 각 문자 입력
   - VK_RETURN으로 Enter 전송

### fpUSpread80이란?

- **Farpoint Spread**: 상용 ActiveX 스프레드시트 컨트롤
- **용도**: Excel과 유사한 그리드 UI
- **사원등록 프로그램**: 3개의 fpUSpread80 컨트롤 사용
  - Spread #0: 용도 불명
  - Spread #1: 용도 불명
  - **Spread #2**: 왼쪽 직원 목록 (우리가 사용하는 것)

## 좌표 매핑

Spread #2의 확립된 좌표:

| 필드 | X 좌표 | Y 좌표 | 설명 |
|------|--------|--------|------|
| 사번 | 50 | 30 | Employee Number |
| 성명 | 100 | 30 | Name |
| 주민번호 | 200 | 30 | ID Number |
| 나이 | 320 | 30 | Age |

**참고**:
- Y=30: 새 행 추가 위치
- 좌표는 컨트롤 내부 상대 좌표 (클라이언트 좌표)
- 화면 절대 좌표가 아님

## 코드 예제

### 기본 사용법

```python
from employee_input import EmployeeInput
from tab_automation import TabAutomation
import time

# 1. 연결
emp_input = EmployeeInput()
emp_input.connect()

# 2. 기본사항 탭 선택
tab_auto = TabAutomation()
tab_auto.connect()
tab_auto.select_tab("기본사항")
time.sleep(0.5)

# 3. 직원 정보 입력
result = emp_input.input_employee(
    employee_no="2025100",
    name="테스트사원",
    id_number="920315-1234567",
    age="33"
)

# 4. 결과 확인
if result['success']:
    print(f"✅ {result['success_count']}/{result['total']}개 입력 완료")
    for item in result['results']:
        if item['success']:
            print(f"  ✓ {item['field']}: {item['value']}")
else:
    print("❌ 입력 실패")
```

### 저수준 입력 (고급)

```python
import win32api
import win32con
import time

def input_to_spread_cell(hwnd, x, y, text):
    """fpUSpread80 셀에 직접 입력"""

    # 1. 셀 클릭하여 선택
    lparam = win32api.MAKELONG(x, y)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN,
                         win32con.MK_LBUTTON, lparam)
    time.sleep(0.03)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(0.2)

    # 2. 각 문자 입력
    for char in text:
        win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
        time.sleep(0.015)

    # 3. Enter 키 전송
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN,
                         win32con.VK_RETURN, 0)
    time.sleep(0.02)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP,
                         win32con.VK_RETURN, 0)
    time.sleep(0.5)

# 사용 예제
spread_hwnd = 0x00080D4C  # Spread #2 핸들
input_to_spread_cell(spread_hwnd, 50, 30, "2025100")  # 사번
```

### Spread 컨트롤 찾기

```python
from pywinauto import application

app = application.Application(backend="win32")
app.connect(title="사원등록")
dlg = app.window(title="사원등록")

# 모든 fpUSpread80 찾기
spread_controls = []
for ctrl in dlg.descendants():
    try:
        if ctrl.class_name() == "fpUSpread80":
            spread_controls.append(ctrl)
    except:
        pass

# Spread #2 = 왼쪽 직원 목록
if len(spread_controls) >= 3:
    left_list = spread_controls[2]
    hwnd = left_list.handle
    print(f"Spread #2 HWND: 0x{hwnd:08X}")
```

## 발견 과정

18회의 시도를 통해 올바른 방법을 확립했습니다:

### Timeline

| Attempt | 내용 | 결과 |
|---------|------|------|
| 08 | SPR32DU80EditHScroll 직접 입력 | 일부 성공 (사번만) |
| 09 | SPR32DU80 + EN_CHANGE 알림 | 보고된 성공, 실제 실패 |
| 10 | visible 컨트롤만 필터링 | 1개만 발견, 목표 아님 |
| 11 | WindowFromPoint 좌표 스캔 | ✨ fpUSpread80 발견! |
| 12 | fpUSpread80 첫 입력 시도 | ✅ "비고" 필드 성공 |
| 13 | fpUSpread80 다중 셀 테스트 | ✅ 10/10 성공 |
| 14 | (번호 사용 안 함) | - |
| 15 | 3개 Spread 전체 테스트 | ✅ 모두 입력 가능 확인 |
| 16 | 좌표 그리드 서치 | 다양한 좌표 테스트 |
| 17 | 왼쪽 목록 컬럼 매핑 | ✅ 4개 필드 좌표 확정 |
| 18 | 완전한 직원 정보 입력 | ✅ 4/4 성공, 방법 확립 |

### 핵심 발견

#### Attempt 11: fpUSpread80 발견

```python
# 좌표 스캔으로 컨트롤 발견
import win32gui

for screen_x in range(100, 1400, 100):
    for screen_y in range(150, 900, 50):
        ctrl_hwnd = win32gui.WindowFromPoint((screen_x, screen_y))
        class_name = win32gui.GetClassName(ctrl_hwnd)
        if "fpUSpread" in class_name:
            print(f"Found: {class_name} at ({screen_x}, {screen_y})")
```

#### Attempt 17: 좌표 매핑

```python
# 체계적인 컬럼 테스트
column_tests = [
    {"x": 50,  "label": "컬럼1(사번추정)", "data": "TEST001"},
    {"x": 100, "label": "컬럼2(성명추정)", "data": "김테스트"},
    {"x": 200, "label": "컬럼3(주민번호추정)", "data": "900101-1234567"},
    {"x": 320, "label": "컬럼4(나이추정)", "data": "35"},
]

# 결과: 모두 row 234에 입력됨 (스크린샷으로 확인)
```

#### Attempt 18: 완전한 통합

```python
# 완전한 직원 데이터
employee_data = {
    "사번": "2025100",
    "성명": "테스트사원",
    "주민번호": "920315-1234567",
    "나이": "33"
}

# 순차 입력: 사번 → 성명 → 주민번호 → 나이
# 결과: 4/4 성공
```

## 실패한 방법들

### ❌ 방법 1: SPR32DU80EditHScroll 사용

**시도한 것**:
```python
# Attempt 08-09
ctrl = edit_controls[0]  # SPR32DU80EditHScroll
win32api.SendMessage(ctrl.handle, win32con.WM_SETTEXT, 0, "2025001")
```

**결과**:
- 사번 필드만 일부 작동
- 성명, 주민번호는 입력 안 됨
- 사용자 확인: "직원 정보 입력 되는 거 맞아? 안 되는 것 같은데?"

**이유**:
- SPR32DU80 컨트롤은 숨겨진 내부 컨트롤 (visible=False)
- 실제 UI와 연결되지 않음
- 단순한 버퍼 역할만 함

### ❌ 방법 2: set_edit_text() 사용

**시도한 것**:
```python
ctrl.set_edit_text("홍길동")
```

**결과**:
- `pywintypes.error: (5, 'SetFocus', '액세스가 거부되었습니다.')`

**이유**:
- pywinauto의 set_edit_text()는 내부적으로 SetFocus 호출
- 사원등록 프로그램이 외부 SetFocus를 차단함

### ❌ 방법 3: SetFocus() 직접 호출

**시도한 것**:
```python
win32gui.SetFocus(hwnd)
win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
```

**결과**:
- `error: (5, 'SetFocus', '액세스가 거부되었습니다.')`

**이유**:
- 동일한 권한 문제

### ❌ 방법 4: WM_SETTEXT만 사용

**시도한 것**:
```python
win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
```

**결과**:
- 내부 버퍼는 변경됨
- UI에 반영 안 됨

**이유**:
- fpUSpread80는 Edit 컨트롤이 아님
- WM_SETTEXT를 지원하지 않음

## 트러블슈팅

### 문제: "fpUSpread80 컨트롤 부족" 에러

```
Exception: fpUSpread80 컨트롤 부족 (발견: 2, 필요: 3)
```

**원인**: 기본사항 탭이 선택되지 않음

**해결**:
```python
from tab_automation import TabAutomation
tab_auto = TabAutomation()
tab_auto.connect()
tab_auto.select_tab("기본사항")
time.sleep(0.5)  # 탭 전환 대기
```

### 문제: 입력은 되지만 엉뚱한 위치에 입력됨

**원인**: 잘못된 Spread 컨트롤 사용

**확인**:
```python
# 반드시 Spread #2 사용
spread_controls = []
for ctrl in dlg.descendants():
    if ctrl.class_name() == "fpUSpread80":
        spread_controls.append(ctrl)

# 인덱스 2 = 세 번째 = 왼쪽 목록
left_list = spread_controls[2]  # ✅ 맞음
wrong_spread = spread_controls[0]  # ❌ 틀림
```

### 문제: 좌표가 작동하지 않음

**원인**: 화면 절대 좌표와 클라이언트 좌표 혼동

**해결**:
- WM_LBUTTONDOWN의 LPARAM은 **클라이언트 좌표** (컨트롤 내부 상대 좌표)
- 화면 절대 좌표가 아님!

```python
# ✅ 맞음: 클라이언트 좌표
lparam = win32api.MAKELONG(50, 30)

# ❌ 틀림: 화면 절대 좌표 사용
screen_x, screen_y = 500, 400
lparam = win32api.MAKELONG(screen_x, screen_y)
```

### 문제: 문자가 누락되거나 순서가 뒤바뀜

**원인**: 타이밍 문제

**해결**:
```python
# 각 문자 사이 충분한 대기
for char in text:
    win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
    time.sleep(0.015)  # 최소 15ms 대기

# 메시지 사이 대기
time.sleep(0.2)  # 클릭 후
time.sleep(0.5)  # Enter 후
```

### 문제: 한글이 입력 안 됨

**원인**: 한글은 WM_CHAR로 직접 입력 가능

**해결**:
```python
# UTF-8 문자열의 각 문자를 ord()로 변환
text = "테스트사원"
for char in text:
    win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
    time.sleep(0.015)

# ord('테') = 53580 (유니코드)
# fpUSpread80가 유니코드 지원
```

## 참고 자료

### 관련 문서
- [탭 자동화](tab-automation.md) - 기본사항 탭 선택 방법
- [윈도우 구조](window-architecture.md) - 사원등록 프로그램 구조
- [테스트 프레임워크](testing-framework.md) - attempt 패턴 설명

### 코드 위치
- `employee_input.py` - 프로덕션 모듈
- `test/attempt/attempt17_left_list_systematic.py` - 좌표 매핑
- `test/attempt/attempt18_complete_employee_input.py` - 완전한 통합 테스트

### 스크린샷
- `test/image/attempt17_02_col1_컬럼1(사번추정).png` - 사번 입력 확인
- `test/image/attempt17_02_col2_컬럼2(성명추정).png` - 성명 입력 확인
- `test/image/attempt18_06_final.png` - 최종 결과 (4개 필드 모두)

## 다음 단계

직원 정보 입력이 확립되었으므로, 다음 단계는:

1. **우측 상세 입력 폼 연구**
   - 우측 폼에도 fpUSpread80가 있을 수 있음
   - 또는 다른 컨트롤 타입 사용 가능

2. **부양가족 데이터 입력**
   - 부양가족 탭으로 전환
   - 유사한 방법 적용

3. **데이터 검증**
   - 입력 후 데이터 읽기
   - 올바르게 저장되었는지 확인

4. **저장 및 확인**
   - 저장 버튼 클릭 자동화
   - 확인 대화상자 처리
