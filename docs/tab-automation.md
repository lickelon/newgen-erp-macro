# 탭 자동화

사원등록 프로그램의 탭을 마우스 커서 움직임 없이 자동으로 선택하는 방법

## ✅ 성공한 방법 (2025-10-30)

### attempt06: 탭 컨트롤 HWND에 직접 WM_LBUTTONDOWN/UP 전송

**핵심**: 탭 컨트롤 HWND를 찾아서 직접 클릭 메시지 전송

**중요**: 클래스명이 프로그램 재시작 시 바뀔 수 있으므로 부분 매칭 사용!

### 완전한 코드 예제

```python
import time
import win32api
import win32con
from pywinauto import application

# 1. 사원등록 윈도우에 연결
app = application.Application(backend="win32")
app.connect(title="사원등록")
dlg = app.window(title="사원등록")

# 2. 탭 컨트롤 찾기 (부분 매칭 - 안정적)
tab_control = None
for ctrl in dlg.descendants():
    if ctrl.class_name().startswith("Afx:TabWnd:"):
        tab_control = ctrl
        break

if tab_control is None:
    raise Exception("탭 컨트롤을 찾을 수 없습니다")

tab_hwnd = tab_control.handle

# 3. 클릭할 탭 위치 설정 (클라이언트 좌표)
tab_positions = {
    "기본사항": (50, 15),
    "부양가족정보": (150, 15),
    "소득자료": (250, 15),
}

tab_name = "부양가족정보"
x, y = tab_positions[tab_name]

# 4. LPARAM 생성
lparam = win32api.MAKELONG(x, y)

# 5. WM_LBUTTONDOWN 전송
win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
time.sleep(0.1)

# 6. WM_LBUTTONUP 전송
win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
time.sleep(0.5)

print(f"✅ '{tab_name}' 탭 선택 완료")
```

### 결과
- ✅ 탭 선택 성공
- ✅ 마우스 커서 움직이지 않음
- ✅ 윈도우 메시지만 사용
- ✅ 백그라운드에서 작동 가능

### 탭 위치 매핑
- **기본사항**: x=50, y=15
- **부양가족정보**: x=150, y=15
- **소득자료**: x=250, y=15

좌표는 **클라이언트 좌표** (탭 컨트롤 내부 기준)

## 사용하기 쉬운 모듈: `tab_automation.py`

```python
from tab_automation import TabAutomation

# 탭 자동화 객체 생성
tab_auto = TabAutomation()

# 연결
tab_auto.connect()

# 탭 선택 (이름으로)
tab_auto.select_tab("부양가족정보")

# 또는 인덱스로
tab_auto.select_tab_by_index(1)  # 0=기본사항, 1=부양가족정보, 2=소득자료
```

## ❌ 실패한 방법들

### 1. TCM_SETCURSEL 메시지 (0x130C) - attempt03
```python
TCM_SETCURSEL = 0x130C
result = win32api.SendMessage(hwnd, TCM_SETCURSEL, tab_index, 0)
# 결과: 0 (실패)
```

**이유**:
- 표준 탭 컨트롤 메시지
- MFC 커스텀 탭에서 작동 안 함
- SendMessage 결과: 0 (실패)

### 2. 부모 윈도우에 WM_LBUTTONDOWN/UP 전송 - attempt03
```python
parent_hwnd = dlg.handle
win32api.SendMessage(parent_hwnd, win32con.WM_LBUTTONDOWN, ...)
```

**이유**:
- 부모 윈도우에 메시지 전송
- 탭 선택 이벤트 발생 안 함
- **탭 컨트롤 자체**가 메시지를 받아야 함

### 3. 탭 컨트롤 자식 요소 클릭 - attempt01
```python
tab_control = dlg.child_window(class_name="Afx:TabWnd:...")
for child in tab_control.descendants():
    child.click_input()
```

**이유**:
- pywinauto의 click_input() 사용
- 46개 자식 요소 클릭 시도
- 탭 변경 안 됨
- **마우스 움직임 발생** (제약사항 위반)

### 4. 키보드 입력 (VK_RIGHT) - attempt05 (미완성)
```python
win32api.SetFocus(hwnd)
win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
```

**이유**:
- 화살표 키로 탭 이동
- 포커스 문제로 작동 안 함

### 5. UIA 백엔드 - attempt04
```python
app_uia = application.Application(backend="uia")
app_uia.connect(title="사원등록")
```

**이유**:
- UI Automation 백엔드 시도
- 탭 컨트롤을 찾을 수 없음
- UIA는 이 MFC 애플리케이션과 호환성 낮음

## 핵심 원리

### 왜 이 방법이 작동하는가?

1. **올바른 타겟**: 탭 컨트롤 HWND에 직접 전송
   - ❌ 부모 윈도우
   - ✅ 탭 컨트롤 자체

2. **올바른 메시지**: WM_LBUTTONDOWN + WM_LBUTTONUP
   - ❌ TCM_SETCURSEL (표준 메시지)
   - ✅ WM_LBUTTONDOWN/UP (클릭 메시지)

3. **올바른 좌표**: 클라이언트 좌표
   - ❌ 화면 절대 좌표
   - ✅ 탭 컨트롤 내부 기준 좌표

4. **올바른 API**: SendMessage
   - ✅ 마우스 커서 움직이지 않음
   - ✅ 메시지만 전달

### 동작 과정

```
1. 탭 컨트롤 HWND 획득
   ↓
2. 클릭 위치 좌표 계산 (클라이언트 좌표)
   ↓
3. WM_LBUTTONDOWN 전송 → 탭 컨트롤이 마우스 다운 감지
   ↓
4. 짧은 대기 (0.1초)
   ↓
5. WM_LBUTTONUP 전송 → 탭 컨트롤이 클릭 완료 처리
   ↓
6. 탭 변경 + WM_NOTIFY로 부모에게 알림
```

## 트러블슈팅

### 탭이 선택되지 않음

**원인 1**: 잘못된 HWND
```python
# ❌ 하드코딩된 HWND
tab_hwnd = 0x000608EE

# ✅ 매번 찾기
for ctrl in dlg.descendants():
    if ctrl.class_name().startswith("Afx:TabWnd:"):
        tab_hwnd = ctrl.handle
```

**원인 2**: 잘못된 클래스명
```python
# ❌ 전체 클래스명 (프로그램 재시작 시 변경)
tab_control = dlg.child_window(class_name="Afx:TabWnd:cd0000:8:10003:10")

# ✅ 부분 매칭
if ctrl.class_name().startswith("Afx:TabWnd:"):
```

**원인 3**: 잘못된 좌표
```python
# ❌ 화면 절대 좌표
x, y = 4405, 1203

# ✅ 클라이언트 좌표 (탭 컨트롤 내부 기준)
x, y = 150, 15
```

### 마우스가 움직임

**원인**: `click_input()` 또는 물리적 클릭 사용
```python
# ❌ 마우스 움직임
child.click_input()

# ✅ SendMessage만 사용
win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, ...)
```

## 참고 자료

- attempt06 코드: `test/attempt/attempt06_direct_tab_hwnd.py`
- 모듈: `tab_automation.py`
- 테스트 스크립트: `test.py`
- 스크린샷: `test/image/attempt06_*.png`
