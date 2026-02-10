# 분납적용 다이얼로그 찾기

**작성일:** 2025-11-20
**총 시도 횟수:** 10회 (attempt 91-100)
**최종 성공:** 같은 프로세스의 최상위 창 열거 ✅

---

## 📋 목차

1. [문제 상황](#문제-상황)
2. [시도한 방법들](#시도한-방법들)
3. [성공한 방법](#성공한-방법)
4. [핵심 발견사항](#핵심-발견사항)
5. [최종 해결책](#최종-해결책)

---

## 문제 상황

### 목표
급여자료입력 프로그램의 "분납적용" 다이얼로그를 찾아서 자동화

### 초기 상태
- 급여자료입력 프로그램에서 "분납적용" 버튼 클릭 시 다이얼로그가 열림
- 스크린샷에서는 명확히 보이지만 pywinauto로 찾을 수 없음
- `descendants()`로 검색해도 발견되지 않음

### 문제점
```python
# ❌ 실패한 방법
all_controls = main_window.descendants()
for ctrl in all_controls:
    if ctrl.class_name() == "#32770" and "분납" in ctrl.window_text():
        # 찾지 못함!
```

---

## 시도한 방법들

### Attempt 91: 모든 창 나열
**목적:** 전체 시스템에서 분납 관련 창 찾기

```python
from pywinauto import findwindows
windows = findwindows.find_elements()
for w in windows:
    if "분납" in w.name:
        print("찾음!")
```

**결과:** ❌ 실패 - 독립 창으로는 찾을 수 없음

---

### Attempt 92: 급여자료입력 자식 다이얼로그 확인
**목적:** 급여자료입력 프로그램의 모든 #32770 다이얼로그 확인

```python
for ctrl in main_window.descendants():
    if ctrl.class_name() == "#32770":
        title = ctrl.window_text()
        if "분납" in title:
            print("찾음!")
```

**결과:** ❌ 실패 - '사원정보', '임금대장'만 발견, 분납적용은 없음

---

### Attempt 93-94: 분납적용 버튼 클릭
**목적:** 버튼을 프로그래밍 방식으로 클릭하여 다이얼로그 열기

```python
# Attempt 93: click_input() - 권한 오류
button.click_input()  # ❌ AccessDenied

# Attempt 94: SendMessage
win32api.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)  # 실행됨
```

**결과:** ❌ 클릭은 실행되었으나 다이얼로그 여전히 못 찾음

---

### Attempt 95: 화면 캡처
**목적:** 현재 상태 확인

```python
capture_window(main_window.handle, "attempt95_salary_window.png")
```

**결과:** ✅ 스크린샷에서 분납적용 다이얼로그 확인됨
- 다이얼로그가 명확히 보임
- 버튼들: '분납적용(Tab)', '취소(Esc)', '인쇄(F9)' 등
- 사원 목록과 입력 영역 존재

---

### Attempt 96: 정확한 다이얼로그 찾기
**목적:** 다양한 방법으로 분납 관련 다이얼로그 찾기

```python
# 방법 1: descendants()에서 #32770 찾기
# 방법 2: Desktop의 최상위 창 확인
# 방법 3: 모든 #32770 출력
```

**결과:** ❌ 여전히 못 찾음

---

### Attempt 97-98: 컨트롤 구조 출력
**목적:** print_control_identifiers로 전체 구조 확인

```python
# Attempt 97: print_control_identifiers()
main_window.print_control_identifiers()  # ❌ 권한 오류

# Attempt 98: 수동으로 descendants() 출력
for ctrl in main_window.descendants():
    print(ctrl.class_name(), ctrl.window_text())
```

**결과:** ❌ 분납적용 버튼만 찾고 다이얼로그는 못 찾음

---

### Attempt 99: 같은 프로세스의 모든 창 찾기 ⭐
**목적:** 급여자료입력과 같은 프로세스의 모든 최상위 창 열거

```python
import win32process
import win32gui

# 프로세스 ID 가져오기
_, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

# 같은 프로세스의 모든 창 찾기
found_windows = []

def enum_callback(hwnd, results):
    if win32gui.IsWindowVisible(hwnd):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid == process_id:
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            results.append((hwnd, title, class_name))
    return True

win32gui.EnumWindows(enum_callback, found_windows)
```

**결과:** ✅ 성공!
- 총 2개 창 발견
- HWND: 0x0026085E, 클래스: #32770, 제목: **(없음)**
- HWND: 0x00030A64, 클래스: Afx:..., 제목: '급여자료입력'

---

### Attempt 100: 제목 없는 다이얼로그 분석 ⭐
**목적:** 발견된 제목 없는 #32770 다이얼로그 상세 분석

```python
dialog_win = app.window(handle=0x0026085E)
children = dialog_win.children()

# 자식 컨트롤 확인
for child in children:
    class_name = child.class_name()
    text = child.window_text()
    print(f"[{class_name}] '{text}'")
```

**결과:** ✅ 완벽한 성공!

**발견된 컨트롤:**
- fpUSpread80: 2개 (사원 목록, 데이터 입력)
- Button: 6개
  - '분납적용(Tab)' ⭐
  - '분납(환급)계산' ⭐
  - '적용해제'
  - '취소(Esc)'
  - '인쇄(F9)'
  - '연말정산불러오기'
- Static: 2개 (안내 텍스트)

---

## 성공한 방법

### 핵심 코드

```python
import win32process
import win32gui
from pywinauto import Application

def find_installment_dialog():
    """
    분납적용 다이얼로그 찾기

    Returns:
        pywinauto 윈도우 객체 또는 None
    """
    # 1. 급여자료입력 프로그램에 연결
    app = Application(backend="win32")
    app.connect(title="급여자료입력")
    main_window = app.window(title="급여자료입력")

    # 2. 프로세스 ID 가져오기
    _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

    # 3. 같은 프로세스의 모든 #32770 창 찾기
    found_dialogs = []

    def enum_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == process_id:
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "#32770":
                    title = win32gui.GetWindowText(hwnd)
                    results.append((hwnd, title))
        return True

    win32gui.EnumWindows(enum_callback, found_dialogs)

    # 4. 제목 없는 다이얼로그 찾기
    for hwnd, title in found_dialogs:
        if not title:  # 제목이 비어있음!
            dialog = app.window(handle=hwnd)

            # 5. 검증: 분납 관련 버튼 확인
            for child in dialog.children():
                try:
                    text = child.window_text()
                    if "분납" in text:
                        print(f"✓ 분납적용 다이얼로그 발견: 0x{hwnd:08X}")
                        return dialog
                except:
                    pass

    return None


# 사용 예시
installment_dlg = find_installment_dialog()
if installment_dlg:
    print("다이얼로그를 찾았습니다!")
    print(f"HWND: 0x{installment_dlg.handle:08X}")
```

---

## 핵심 발견사항

### 1. 다이얼로그 제목이 빈 문자열 ⚠️

**문제:**
- `ctrl.window_text()`가 빈 문자열 `""` 반환
- `if "분납" in title` 조건으로 찾을 수 없음

**이유:**
- MFC 다이얼로그가 제목 없이 생성됨
- 시각적으로는 "분납적용"이라는 제목이 보이지만 실제 윈도우 제목은 비어있음

### 2. descendants()로 찾을 수 없음

**문제:**
- `main_window.descendants()`에 포함되지 않음
- 급여자료입력 창의 자식이 아님

**이유:**
- 분납적용 다이얼로그가 독립적인 최상위 창으로 생성됨
- 같은 프로세스이지만 부모-자식 관계가 아님

### 3. 같은 프로세스의 최상위 창으로 존재

**발견:**
```
PID: 3936
├── 0x00030A64 (Afx:...) - "급여자료입력"
└── 0x0026085E (#32770) - ""  ← 분납적용 다이얼로그
```

**특징:**
- 두 창 모두 같은 프로세스에 속함
- 두 창 모두 최상위 창 (top-level window)
- 부모-자식 관계가 아닌 형제 관계

### 4. win32gui.EnumWindows() 필수

**이유:**
- pywinauto의 `descendants()`는 자식 컨트롤만 탐색
- 형제 창을 찾으려면 시스템 레벨 API 필요
- `win32gui.EnumWindows()`로 모든 최상위 창 열거 후 프로세스 ID로 필터링

---

## 최종 해결책

### 프로덕션 코드

```python
"""
분납적용 다이얼로그 자동화 모듈
"""
import win32process
import win32gui
from pywinauto import Application


class InstallmentDialog:
    """분납적용 다이얼로그 자동화"""

    def __init__(self):
        self.app = None
        self.main_window = None
        self.dialog = None
        self.process_id = None

    def connect(self):
        """급여자료입력 프로그램에 연결"""
        self.app = Application(backend="win32")
        self.app.connect(title="급여자료입력")
        self.main_window = self.app.window(title="급여자료입력")

        _, self.process_id = win32process.GetWindowThreadProcessId(
            self.main_window.handle
        )

        print(f"✓ 연결 성공 (PID: {self.process_id})")

    def find_dialog(self):
        """
        분납적용 다이얼로그 찾기

        Returns:
            bool: 찾았으면 True
        """
        if not self.process_id:
            raise RuntimeError("connect()를 먼저 호출하세요")

        # 같은 프로세스의 모든 #32770 찾기
        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == self.process_id:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "#32770":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        # 제목 없는 다이얼로그에서 분납 관련 버튼 확인
        for hwnd, title in found_dialogs:
            if not title:  # 제목이 비어있음
                dialog = self.app.window(handle=hwnd)

                # 검증: 분납 관련 버튼 확인
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text or "분납(환급)계산" in text:
                            self.dialog = dialog
                            print(f"✓ 분납적용 다이얼로그 발견: 0x{hwnd:08X}")
                            return True
                    except:
                        pass

        print("✗ 분납적용 다이얼로그를 찾지 못했습니다")
        return False

    def get_spreads(self):
        """
        스프레드 컨트롤 가져오기

        Returns:
            tuple: (왼쪽 스프레드, 오른쪽 스프레드)
        """
        if not self.dialog:
            raise RuntimeError("find_dialog()를 먼저 호출하세요")

        spreads = []
        for child in self.dialog.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        if len(spreads) < 2:
            raise RuntimeError(f"스프레드를 충분히 찾지 못했습니다 ({len(spreads)}개)")

        # 왼쪽/오른쪽 구분 (X 좌표 기준)
        spreads.sort(key=lambda s: s.rectangle().left)

        return spreads[0], spreads[1]


# 사용 예시
if __name__ == "__main__":
    installer = InstallmentDialog()

    # 1. 연결
    installer.connect()

    # 2. 다이얼로그 찾기
    if installer.find_dialog():
        # 3. 스프레드 가져오기
        left_spread, right_spread = installer.get_spreads()
        print(f"왼쪽 스프레드: 0x{left_spread.handle:08X}")
        print(f"오른쪽 스프레드: 0x{right_spread.handle:08X}")
```

---

## 참고 사항

### 유사한 경우에 적용

다음과 같은 상황에서 이 방법을 사용:

1. **제목 없는 다이얼로그**
   - `window_text()`가 빈 문자열
   - 시각적 제목과 실제 윈도우 제목이 다름

2. **descendants()로 찾을 수 없는 창**
   - 부모-자식 관계가 아님
   - 같은 프로세스의 형제 창

3. **MFC 모달 다이얼로그**
   - `CDialog`로 생성된 다이얼로그
   - 독립적인 최상위 창으로 생성됨

### 디버깅 팁

```python
# 같은 프로세스의 모든 창 확인
def list_process_windows(process_id):
    found = []
    def callback(hwnd, results):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid == process_id:
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            results.append((hwnd, title, class_name))
        return True

    win32gui.EnumWindows(callback, found)

    for hwnd, title, class_name in found:
        print(f"0x{hwnd:08X} | {class_name:30s} | '{title}'")
```

---

## 통계

- **총 시도**: 10회 (attempt 91-100)
- **성공률**: 20% (2/10 - attempt 99, 100)
- **소요 시간**: 약 30분
- **최종 결과**: 분납적용 다이얼로그 발견 및 구조 파악 완료 ✅

---

**작성자:** Claude Code
**테스트 환경:** Windows 11, Python 3.14, pywinauto
**대상 프로그램:** 케이렙 365 - 급여자료입력
**최종 업데이트:** 2025-11-20
