# 사원등록 자동화 완벽 가이드

**작성일:** 2025-10-30
**총 시도 횟수:** 53회
**최종 성공:** 좌표 없는 완전 자동화 달성 ✅

---

## 📋 목차

1. [성공한 방법 요약](#성공한-방법-요약)
2. [사원 선택 (좌표 없음)](#사원-선택-좌표-없음)
3. [탭 전환 (좌표 없음)](#탭-전환-좌표-없음)
4. [데이터 입력 (좌표 없음)](#데이터-입력-좌표-없음)
5. [완전한 자동화 예제](#완전한-자동화-예제)
6. [중요 발견사항](#중요-발견사항)
7. [실패한 방법들](#실패한-방법들)
8. [트러블슈팅](#트러블슈팅)

---

## 성공한 방법 요약

| 작업 | 방법 | Attempt | 좌표 사용 |
|------|------|---------|-----------|
| 사원 선택 | `set_focus()` + 키보드 | 53 | ❌ 없음 |
| 탭 전환 | `ShowWindow(SW_HIDE/SHOW)` | 52 | ❌ 없음 |
| 데이터 입력 | `dlg.type_keys()` | 43 | ❌ 없음 |

**✅ 완전히 좌표 독립적인 자동화 달성!**

---

## 사원 선택 (좌표 없음)

### 문제
- 기존: 좌표 클릭으로 사원 선택
- 해상도나 창 크기 변경 시 작동 불가

### 해결책
**왼쪽 스프레드시트에 포커스 후 키보드로 이동**

### 코드

```python
def select_employee_by_index(dlg, index=0):
    """
    사원을 인덱스로 선택 (좌표 없음)

    Args:
        dlg: pywinauto 윈도우 객체
        index: 사원 인덱스 (0=첫 번째)

    Returns:
        왼쪽 스프레드 객체
    """
    import time

    # 1. 모든 스프레드 찾기
    spreads = []
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "fpUSpread80":
                spreads.append(ctrl)
        except:
            pass

    # 2. 왼쪽 스프레드 = 가장 왼쪽 X 좌표
    spreads.sort(key=lambda s: s.rectangle().left)
    left_spread = spreads[0]  # 사원 목록

    print(f"왼쪽 스프레드: 0x{left_spread.handle:08X}")

    # 3. 포커스 설정
    left_spread.set_focus()
    time.sleep(0.5)

    # 4. Home으로 첫 번째 이동
    dlg.type_keys("{HOME}", pause=0.1)
    time.sleep(0.3)

    # 5. Down 키로 원하는 위치로 이동
    for i in range(index):
        dlg.type_keys("{DOWN}", pause=0.1)
        time.sleep(0.2)

    print(f"사원 인덱스 {index} 선택 완료")
    return left_spread


# 사용 예시
from pywinauto import application

app = application.Application(backend="win32")
app.connect(title="사원등록")
dlg = app.window(title="사원등록")

# 첫 번째 사원 선택
select_employee_by_index(dlg, 0)

# 두 번째 사원 선택
select_employee_by_index(dlg, 1)

# 세 번째 사원 선택
select_employee_by_index(dlg, 2)
```

### 주요 키

| 키 | 기능 |
|----|------|
| `{HOME}` | 첫 번째 사원 |
| `{DOWN}` | 다음 사원 |
| `{UP}` | 이전 사원 |
| `{PGDN}` | 페이지 다운 |
| `{PGUP}` | 페이지 업 |
| `^{HOME}` | Ctrl+Home (맨 처음) |

### 특징
- ✅ 해상도 독립적
- ✅ 창 크기 독립적
- ✅ 안정적으로 작동
- ✅ 사원 인덱스로 선택 가능

---

## 탭 전환 (좌표 없음)

### 문제
- 기존: 좌표 기반 탭 클릭
- 버튼 클릭, TCM_SETCURSEL 등 모두 실패
- 47~50번 시도 모두 실패

### 해결책
**다이얼로그를 ShowWindow로 숨기고/보이기**

### 코드

```python
def switch_tab_by_name(dlg, tab_name):
    """
    탭을 이름으로 전환 (좌표 없음)

    Args:
        dlg: pywinauto 윈도우 객체
        tab_name: 탭 이름 ("기본사항", "부양가족명세", "추가사항")

    Returns:
        전환된 다이얼로그 객체
    """
    import time
    import win32gui
    import win32con

    # 1. 모든 탭 다이얼로그 찾기
    dialogs = {}
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "#32770":
                text = ctrl.window_text().strip()
                if "기본사항" in text:
                    dialogs["기본사항"] = ctrl
                elif "부양가족명세" in text:
                    dialogs["부양가족명세"] = ctrl
                elif "추가사항" in text:
                    dialogs["추가사항"] = ctrl
        except:
            pass

    if tab_name not in dialogs:
        print(f"✗ '{tab_name}' 탭을 찾을 수 없음")
        return None

    target_dialog = dialogs[tab_name]

    # 2. 다른 다이얼로그 모두 숨기기
    for name, dialog in dialogs.items():
        if name != tab_name:
            win32gui.ShowWindow(dialog.handle, win32con.SW_HIDE)

    time.sleep(0.3)

    # 3. 목표 다이얼로그 보이기
    win32gui.ShowWindow(target_dialog.handle, win32con.SW_SHOW)
    time.sleep(0.5)

    print(f"✓ '{tab_name}' 탭으로 전환 완료")
    return target_dialog


# 사용 예시
# 기본사항 탭으로 전환
switch_tab_by_name(dlg, "기본사항")

# 부양가족명세 탭으로 전환
switch_tab_by_name(dlg, "부양가족명세")

# 추가사항 탭으로 전환
switch_tab_by_name(dlg, "추가사항")
```

### 탭 구조

```
사원등록 윈도우
├── Afx:TabWnd (탭 컨트롤)
│   ├── #32770 "   기본사항   " (Dialog)
│   ├── #32770 " 부양가족명세 " (Dialog)
│   └── #32770 "   추가사항   " (Dialog)
```

### 핵심 원리
1. 각 탭은 별도의 `#32770` Dialog
2. `ShowWindow(SW_HIDE)` - 다이얼로그 숨김
3. `ShowWindow(SW_SHOW)` - 다이얼로그 표시
4. 하나만 보이면 해당 탭으로 전환됨

### 특징
- ✅ 좌표 사용 안 함
- ✅ 해상도 독립적
- ✅ 안정적으로 작동
- ✅ 간단한 코드

---

## 데이터 입력 (좌표 없음)

### 문제
- `SendMessage(WM_CHAR)` 실패
- `SendInput` 실패
- 클립보드 붙여넣기 실패
- 39~42번 시도 모두 실패

### 해결책
**pywinauto의 `dlg.type_keys()` 메서드**

### 코드

```python
def input_dependent_data(dlg, data):
    """
    부양가족 데이터 입력 (좌표 없음)

    Args:
        dlg: pywinauto 윈도우 객체
        data: [연말관계, 성명, 내외국, 년도]

    Example:
        data = ["4", "김자녀", "내", "2020"]
    """
    import time
    import win32api
    import win32con

    # 1. 부양가족명세 탭으로 전환
    switch_tab_by_name(dlg, "부양가족명세")

    # 2. 스프레드 찾기
    spread = None
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "fpUSpread80":
                # 오른쪽 스프레드 찾기 (더 큰 너비)
                if spread is None or ctrl.rectangle().width() > spread.rectangle().width():
                    spread = ctrl
        except:
            pass

    if not spread:
        print("✗ 스프레드를 찾지 못함")
        return False

    print(f"스프레드: 0x{spread.handle:08X}")

    # 3. Down 키로 입력 행 이동
    hwnd = spread.handle
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
    time.sleep(0.02)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
    time.sleep(0.5)

    # 4. 데이터 입력
    field_names = ["연말관계", "성명", "내외국", "년도"]

    for idx, value in enumerate(data):
        print(f"  [{idx+1}/4] {field_names[idx]}: \"{value}\"")

        # dlg.type_keys() 사용!
        dlg.type_keys(value, with_spaces=False, pause=0.05)
        time.sleep(0.3)

        # Tab (마지막 필드 제외)
        if idx < len(data) - 1:
            dlg.type_keys("{TAB}", pause=0.05)
            time.sleep(0.2)

    # 5. Enter로 확정
    dlg.type_keys("{ENTER}", pause=0.05)
    time.sleep(0.5)

    print("✓ 데이터 입력 완료")
    return True


# 사용 예시
data = [
    "4",      # 연말관계: 4=자녀
    "김자녀",  # 성명
    "내",      # 내/외국인
    "2020"    # 출생년도
]

input_dependent_data(dlg, data)
```

### 연말관계 코드

| 코드 | 관계 |
|------|------|
| 0 | 본인 |
| 1 | 부모 |
| 2 | 배우자부모 |
| 3 | 배우자 |
| 4 | 자녀 |

### 핵심 포인트
- ✅ `dlg.type_keys()` 사용 (spread.type_keys() 아님!)
- ✅ `with_spaces=False` 옵션
- ✅ `pause=0.05` 적절한 딜레이

### 특징
- ✅ 좌표 사용 안 함
- ✅ 한글 입력 완벽 지원
- ✅ 안정적으로 작동

---

## 완전한 자동화 예제

### 통합 자동화 함수

```python
from pywinauto import application
import time
import win32api
import win32con
import win32gui


def select_employee_by_index(dlg, index):
    """사원 선택 (좌표 없음)"""
    spreads = []
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "fpUSpread80":
                spreads.append(ctrl)
        except:
            pass

    spreads.sort(key=lambda s: s.rectangle().left)
    left_spread = spreads[0]

    left_spread.set_focus()
    time.sleep(0.5)

    dlg.type_keys("{HOME}", pause=0.1)
    time.sleep(0.3)

    for i in range(index):
        dlg.type_keys("{DOWN}", pause=0.1)
        time.sleep(0.2)

    return left_spread


def switch_tab_by_name(dlg, tab_name):
    """탭 전환 (좌표 없음)"""
    dialogs = {}
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "#32770":
                text = ctrl.window_text().strip()
                if "기본사항" in text:
                    dialogs["기본사항"] = ctrl
                elif "부양가족명세" in text:
                    dialogs["부양가족명세"] = ctrl
                elif "추가사항" in text:
                    dialogs["추가사항"] = ctrl
        except:
            pass

    if tab_name not in dialogs:
        return None

    target_dialog = dialogs[tab_name]

    for name, dialog in dialogs.items():
        if name != tab_name:
            win32gui.ShowWindow(dialog.handle, win32con.SW_HIDE)

    time.sleep(0.3)
    win32gui.ShowWindow(target_dialog.handle, win32con.SW_SHOW)
    time.sleep(0.5)

    return target_dialog


def input_dependent_data(dlg, data):
    """부양가족 데이터 입력 (좌표 없음)"""
    switch_tab_by_name(dlg, "부양가족명세")

    spread = None
    spreads = []
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "fpUSpread80":
                spreads.append(ctrl)
        except:
            pass

    # 오른쪽 스프레드 = 더 넓은 것
    spreads.sort(key=lambda s: s.rectangle().width(), reverse=True)
    spread = spreads[0]

    hwnd = spread.handle
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
    time.sleep(0.02)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
    time.sleep(0.5)

    for idx, value in enumerate(data):
        dlg.type_keys(value, with_spaces=False, pause=0.05)
        time.sleep(0.3)

        if idx < len(data) - 1:
            dlg.type_keys("{TAB}", pause=0.05)
            time.sleep(0.2)

    dlg.type_keys("{ENTER}", pause=0.05)
    time.sleep(0.5)

    return True


# 메인 자동화
def automate_employee_dependents(employee_index, dependents_list):
    """
    사원의 부양가족 일괄 등록

    Args:
        employee_index: 사원 인덱스
        dependents_list: 부양가족 데이터 리스트

    Example:
        automate_employee_dependents(1, [
            ["4", "김자녀1", "내", "2020"],
            ["4", "김자녀2", "내", "2022"],
            ["3", "이배우자", "내", "1990"]
        ])
    """
    # 연결
    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    print(f"\n{'='*60}")
    print(f"사원 인덱스 {employee_index} 부양가족 등록 시작")
    print(f"{'='*60}")

    # 1. 사원 선택
    print(f"\n[1/3] 사원 선택 중...")
    select_employee_by_index(dlg, employee_index)
    print(f"  ✓ 사원 인덱스 {employee_index} 선택 완료")

    # 2. 부양가족명세 탭으로 전환
    print(f"\n[2/3] 부양가족명세 탭으로 전환 중...")
    switch_tab_by_name(dlg, "부양가족명세")
    print(f"  ✓ 탭 전환 완료")

    # 3. 부양가족 입력
    print(f"\n[3/3] 부양가족 {len(dependents_list)}명 입력 중...")
    for idx, dependent_data in enumerate(dependents_list):
        print(f"\n  [{idx+1}/{len(dependents_list)}] {dependent_data[1]} 입력 중...")
        success = input_dependent_data(dlg, dependent_data)
        if success:
            print(f"    ✓ 성공")
        else:
            print(f"    ✗ 실패")

    print(f"\n{'='*60}")
    print(f"완료")
    print(f"{'='*60}")


# 실행 예시
if __name__ == "__main__":
    # 두 번째 사원의 부양가족 3명 등록
    automate_employee_dependents(1, [
        ["4", "김자녀", "내", "2020"],
        ["3", "이배우자", "내", "1995"],
        ["1", "박부모", "내", "1965"]
    ])
```

---

## 중요 발견사항

### 1. 관리자 권한 필수 ⚠️

**문제:**
- 일반 권한에서는 탭 전환, 입력 모두 실패
- `SendMessage`, `ShowWindow` 등이 무시됨

**원인:**
- 사원등록 프로그램이 관리자 권한으로 실행
- Windows UIPI (User Interface Privilege Isolation) 보안

**해결:**
```bash
# 관리자 권한 터미널에서 실행
# PowerShell 관리자 모드로 열기 → 실행
uv run python your_script.py
```

### 2. fpUSpread80 컨트롤 특성

**발견:**
- Farpoint Spread ActiveX 컨트롤
- `spread.type_keys()` ❌ 작동 안 함 (ElementNotVisible)
- `dlg.type_keys()` ✅ 작동함!

**이유:**
- 스프레드가 visible=False로 판단됨
- 다이얼로그 레벨에서 키 입력해야 함

### 3. 다이얼로그 구조

```
사원등록 (Afx:FrameOrView:4...)
├── Afx:TabWnd:* (탭 컨트롤)
│   ├── Button (텍스트 없음) - 클릭 불가 ❌
│   ├── Button (텍스트 없음) - 클릭 불가 ❌
│   ├── Button (텍스트 없음) - 클릭 불가 ❌
│   ├── #32770 "   기본사항   "
│   │   └── (폼 필드들...)
│   ├── #32770 " 부양가족명세 "
│   │   └── fpUSpread80 (오른쪽 스프레드)
│   └── #32770 "   추가사항   "
│       └── (폼 필드들...)
└── fpUSpread80 (왼쪽 사원 목록 스프레드)
```

### 4. 스프레드 구분법

```python
# 모든 스프레드 찾기
spreads = []
for ctrl in dlg.descendants():
    if ctrl.class_name() == "fpUSpread80":
        spreads.append(ctrl)

# 왼쪽 스프레드 (사원 목록) = X 좌표가 작음
spreads.sort(key=lambda s: s.rectangle().left)
left_spread = spreads[0]

# 오른쪽 스프레드 (부양가족) = 너비가 넓음
spreads.sort(key=lambda s: s.rectangle().width(), reverse=True)
right_spread = spreads[0]
```

### 5. 타이밍 권장값

| 작업 | 대기 시간 |
|------|-----------|
| `set_focus()` 후 | 0.5초 |
| `ShowWindow()` 후 | 0.3~0.5초 |
| Down 키 후 | 0.2~0.5초 |
| 필드 입력 후 | 0.3초 |
| Tab 키 후 | 0.2초 |
| Enter 키 후 | 0.5초 |

**너무 짧으면:**
- 입력이 누락될 수 있음
- 탭 전환이 완료 안 됨

**너무 길면:**
- 자동화 속도가 느려짐

---

## 실패한 방법들

### 탭 선택 실패 (Attempt 47-51)

| Attempt | 방법 | 결과 |
|---------|------|------|
| 47 | 페이지 텍스트로 버튼 찾아 클릭 | ❌ 클릭해도 탭 안 바뀜 |
| 48 | BM_CLICK 메시지 | ❌ 탭 안 바뀜 |
| 49 | pywinauto `click()` | ❌ 탭 안 바뀜 |
| 50 | TCM_SETCURSEL 메시지 | ❌ 인덱스 변경 안 됨 |
| 51 | `set_focus()` 만 | ❌ 탭 안 바뀜 |

**결론:** 버튼은 장식용, 실제 탭은 Dialog의 show/hide로 제어됨

### 데이터 입력 실패 (Attempt 34-42)

| Attempt | 방법 | 결과 |
|---------|------|------|
| 34-36 | WM_CHAR + SendMessage | ❌ 입력 안 됨 |
| 39 | 클립보드 + Ctrl+V | ❌ 입력 안 됨 |
| 40 | SendInput (전역) | ❌ 입력 안 됨 |
| 41 | SetForegroundWindow + SendInput | ❌ 입력 안 됨 |
| 42 | `spread.type_keys()` | ❌ ElementNotVisible |

**결론:** `dlg.type_keys()`만 작동함

---

## 트러블슈팅

### 문제: 탭이 전환되지 않음

**증상:**
```python
switch_tab_by_name(dlg, "부양가족명세")
# 탭이 바뀌지 않음
```

**원인:**
1. 관리자 권한 없음
2. 대기 시간 부족
3. 다이얼로그를 찾지 못함

**해결:**
```python
# 1. 관리자 권한 확인
# PowerShell을 관리자로 실행

# 2. 대기 시간 늘리기
time.sleep(0.5)  # 0.3 → 0.5

# 3. 디버깅: 다이얼로그 출력
for ctrl in dlg.descendants():
    if ctrl.class_name() == "#32770":
        print(f"Dialog: {ctrl.window_text()}")
```

### 문제: 데이터가 입력되지 않음

**증상:**
```python
dlg.type_keys("테스트", ...)
# 아무 것도 입력 안 됨
```

**원인:**
1. 관리자 권한 없음
2. 스프레드에 포커스 없음
3. Down 키를 안 눌러서 입력 행이 아님

**해결:**
```python
# 1. 관리자 권한 확인

# 2. Down 키 먼저
win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
time.sleep(0.5)

# 3. 그 다음 입력
dlg.type_keys(value, ...)
```

### 문제: 사원을 선택할 수 없음

**증상:**
```python
select_employee_by_index(dlg, 1)
# 사원이 선택 안 됨
```

**원인:**
1. 왼쪽 스프레드를 못 찾음
2. 포커스 설정 실패

**해결:**
```python
# 디버깅: 스프레드 확인
spreads = []
for ctrl in dlg.descendants():
    if ctrl.class_name() == "fpUSpread80":
        rect = ctrl.rectangle()
        print(f"Spread: L={rect.left}, W={rect.width()}")
        spreads.append(ctrl)

# 왼쪽 것 = X 좌표 작은 것
spreads.sort(key=lambda s: s.rectangle().left)
left_spread = spreads[0]
```

### 문제: 32비트/64비트 경고

**증상:**
```
UserWarning: 32-bit application should be automated using 32-bit Python (you use 64-bit Python)
```

**영향:**
- 대부분의 경우 작동함
- 일부 기능에서 문제 발생 가능

**해결:**
```bash
# 32비트 Python 설치 후 사용
# 또는 경고 무시하고 사용 (대부분 문제없음)
```

---

## 부록: 전체 시도 이력

### 성공한 시도

| Attempt | 내용 | 성과 |
|---------|------|------|
| 15 | 좌표 클릭 + WM_CHAR | ✅ 입력 성공 (좌표 의존) |
| 43 | dlg.type_keys() | ✅ 입력 성공 (좌표 독립) |
| 52 | ShowWindow(HIDE/SHOW) | ✅ 탭 전환 성공 |
| 53 | set_focus() + 키보드 | ✅ 사원 선택 성공 |

### 주요 실패 시도

| Attempt 범위 | 내용 | 결과 |
|--------------|------|------|
| 8-18 | SPR32DU80EditHScroll 대상 | ❌ 보이지만 입력 안 됨 |
| 19-33 | 키보드 전용 시도 | ❌ 모두 실패 |
| 34-42 | 다양한 입력 방식 | ❌ 모두 실패 |
| 47-51 | 다양한 탭 선택 방식 | ❌ 모두 실패 |

---

## 최종 요약

### ✅ 달성한 것

1. **사원 선택**: `set_focus()` + 키보드 (좌표 X)
2. **탭 전환**: `ShowWindow()` (좌표 X)
3. **데이터 입력**: `dlg.type_keys()` (좌표 X)
4. **완전한 자동화**: 모든 작업이 좌표 독립적

### 🔑 핵심 교훈

1. **MFC 커스텀 컨트롤**은 표준 방법이 안 통함
2. **pywinauto 고수준 메서드**가 가장 안정적
3. **Dialog 구조 이해**가 핵심
4. **관리자 권한** 필수
5. **적절한 타이밍**이 중요

### 📊 통계

- **총 시도**: 53회
- **성공률**: 7.5% (4/53)
- **소요 시간**: 약 8시간 (추정)
- **최종 결과**: 완전한 좌표 독립적 자동화 달성 ✅

---

**작성자:** Claude Code
**테스트 환경:** Windows 11, Python 3.12, pywinauto 0.6.8
**대상 프로그램:** 케이렙 365 - 사원등록
**최종 업데이트:** 2025-10-30
