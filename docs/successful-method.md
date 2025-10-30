# 사원등록 부양가족정보 입력 자동화 - 성공한 방법

## 최종 성공 결과

**날짜:** 2025-10-30

**총 시도 횟수:** 51회

**성공한 방법:**
1. **탭 선택**: `set_focus()` 메서드로 다이얼로그 직접 활성화 (Attempt 51) ✅
2. **데이터 입력**: `dlg.type_keys()` 메서드 (Attempt 43) ✅

---

## 1. 탭 선택 방법 (좌표 없음) - Attempt 51

### 문제점
- 기존 `TabAutomation`은 좌표 기반 (해상도 의존적)
- 버튼 클릭, 메시지 전송 등 모든 방법 실패 (Attempt 47-50)

### 해결책
**다이얼로그를 직접 찾아서 `set_focus()` 호출**

### 코드

```python
def select_tab_by_dialog(dlg, tab_name):
    """
    다이얼로그를 찾아서 set_focus()로 탭 선택

    Args:
        dlg: pywinauto 윈도우 객체
        tab_name: 탭 이름 (예: "부양가족명세", "기본사항")

    Returns:
        선택된 다이얼로그 객체 또는 None
    """
    import time

    # descendants()로 모든 자식 컨트롤 검색
    for ctrl in dlg.descendants():
        try:
            # #32770 클래스 (Dialog)이고 탭 이름이 포함된 경우
            if ctrl.class_name() == "#32770" and tab_name in ctrl.window_text():
                print(f"✓ '{tab_name}' 다이얼로그 찾음")
                print(f"  HWND: 0x{ctrl.handle:08X}")

                # set_focus()로 탭 전환
                ctrl.set_focus()
                time.sleep(0.5)  # 전환 대기

                return ctrl
        except:
            pass

    print(f"✗ '{tab_name}' 다이얼로그를 찾지 못함")
    return None


# 사용 예시
from pywinauto import application

app = application.Application(backend="win32")
app.connect(title="사원등록")
dlg = app.window(title="사원등록")

# 기본사항 탭 선택
select_tab_by_dialog(dlg, "기본사항")

# 부양가족명세 탭 선택
select_tab_by_dialog(dlg, "부양가족명세")
```

### 핵심 원리
1. `#32770` 클래스는 MFC Dialog를 의미
2. 각 탭은 별도의 Dialog로 구현됨:
   - "기본사항" Dialog
   - "부양가족명세" Dialog
   - "추가사항" Dialog
3. `set_focus()`를 호출하면 해당 Dialog가 활성화되며 **탭이 자동으로 전환됨**

### 장점
- ✅ **좌표 사용 안 함** (해상도 독립적)
- ✅ 간단한 코드
- ✅ 안정적으로 작동

---

## 2. 데이터 입력 방법 - Attempt 43

### 문제점
- `SendMessage(WM_CHAR)` 실패
- `SendInput` 실패
- 클립보드 붙여넣기 실패

### 해결책
**`dlg.type_keys()` 메서드 사용**

### 코드

```python
def input_dependent_data(dlg, data):
    """
    부양가족 데이터 입력

    Args:
        dlg: pywinauto 윈도우 객체
        data: 입력할 데이터 리스트 [연말관계, 성명, 내외국, 년도]

    Example:
        data = ["4", "김자녀", "내", "2020"]
        # 4 = 자녀, 성명, 내국인, 출생년도
    """
    import time

    # 1. 부양가족명세 탭 선택
    select_tab_by_dialog(dlg, "부양가족명세")

    # 2. 스프레드 찾기
    spread = None
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "fpUSpread80":
                spread = ctrl
                break
        except:
            pass

    if not spread:
        print("✗ 스프레드를 찾지 못함")
        return False

    # 3. Down 키로 입력 행 이동
    import win32api
    import win32con

    hwnd = spread.handle
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
    time.sleep(0.02)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
    time.sleep(0.5)

    # 4. 데이터 입력 (dlg.type_keys 사용!)
    field_names = ["연말관계", "성명", "내외국", "년도"]

    for idx, value in enumerate(data):
        print(f"  [{idx+1}/4] {field_names[idx]}: \"{value}\"")

        # 값 입력
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
    "내",      # 내/외국: 내국인
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

---

## 3. 완전한 자동화 예제

```python
from pywinauto import application
import time
import win32api
import win32con


def select_tab_by_dialog(dlg, tab_name):
    """탭 선택"""
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "#32770" and tab_name in ctrl.window_text():
                ctrl.set_focus()
                time.sleep(0.5)
                return ctrl
        except:
            pass
    return None


def input_dependent_data(dlg, data):
    """부양가족 데이터 입력"""
    # 1. 탭 선택
    select_tab_by_dialog(dlg, "부양가족명세")

    # 2. 스프레드 찾기
    spread = None
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "fpUSpread80":
                spread = ctrl
                break
        except:
            pass

    if not spread:
        return False

    # 3. Down 키
    hwnd = spread.handle
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
    time.sleep(0.02)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
    time.sleep(0.5)

    # 4. 데이터 입력
    for idx, value in enumerate(data):
        dlg.type_keys(value, with_spaces=False, pause=0.05)
        time.sleep(0.3)

        if idx < len(data) - 1:
            dlg.type_keys("{TAB}", pause=0.05)
            time.sleep(0.2)

    # 5. Enter
    dlg.type_keys("{ENTER}", pause=0.05)
    time.sleep(0.5)

    return True


# 메인 실행
if __name__ == "__main__":
    # 연결
    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    # 부양가족 데이터
    dependents = [
        ["4", "김자녀", "내", "2020"],
        ["3", "이배우자", "내", "1995"],
        ["1", "박부모", "내", "1960"],
    ]

    # 입력
    for idx, data in enumerate(dependents):
        print(f"\n[{idx+1}/{len(dependents)}] 입력 중...")
        success = input_dependent_data(dlg, data)
        if success:
            print(f"  ✓ 성공: {data}")
        else:
            print(f"  ✗ 실패: {data}")
```

---

## 4. 중요 발견사항

### 관리자 권한 필수
**문제:** 일반 권한에서는 탭 선택 안 됨
**원인:** UIPI (User Interface Privilege Isolation) 보안
**해결:** 관리자 권한 터미널에서 실행

```bash
# 관리자 권한으로 실행
uv run python your_script.py
```

### fpUSpread80 특성
- Farpoint Spread ActiveX 스프레드시트 컨트롤
- 직접 HWND에 메시지 전송 불가
- `dlg.type_keys()`가 유일하게 작동하는 방법

### 탭 구조
```
Afx:TabWnd (탭 컨트롤)
├── Button (텍스트 없음) - 클릭 불가
├── Button (텍스트 없음) - 클릭 불가
├── Button (텍스트 없음) - 클릭 불가
├── #32770 "기본사항" (Dialog)
├── #32770 "부양가족명세" (Dialog) ← 이것에 set_focus()!
└── #32770 "추가사항" (Dialog)
```

---

## 5. 실패한 방법들 (참고용)

| Attempt | 방법 | 결과 |
|---------|------|------|
| 15 | 좌표 클릭 + WM_CHAR | ✅ 성공 (좌표 의존적) |
| 34-36 | WM_CHAR + SendMessage | ❌ 실패 |
| 39 | 클립보드 붙여넣기 | ❌ 실패 |
| 40 | SendInput (전역) | ❌ 실패 |
| 41 | SetForegroundWindow + SendInput | ❌ 실패 |
| 42 | spread.type_keys() | ❌ ElementNotVisible |
| 43 | dlg.type_keys() | ✅ 성공 (입력만) |
| 47 | 페이지 텍스트로 버튼 클릭 | ❌ 실패 |
| 48 | BM_CLICK 메시지 | ❌ 실패 |
| 49 | pywinauto click() | ❌ 실패 |
| 50 | TCM_SETCURSEL 메시지 | ❌ 실패 |
| 51 | Dialog set_focus() | ✅ 성공 (탭 선택) |

---

## 6. 권장사항

### 필수 조건
1. ✅ 관리자 권한으로 실행
2. ✅ pywinauto win32 백엔드 사용
3. ✅ 적절한 대기 시간 (0.3~0.5초)

### 타이밍
```python
# 권장 대기 시간
set_focus() 후: 0.5초
Down 키 후: 0.5초
각 필드 입력 후: 0.3초
Tab 키 후: 0.2초
Enter 키 후: 0.5초
```

### 에러 처리
```python
try:
    result = input_dependent_data(dlg, data)
    if not result:
        print("입력 실패 - 재시도")
except Exception as e:
    print(f"오류 발생: {e}")
```

---

## 7. 결론

**최종 솔루션:**
- 탭 선택: `#32770` Dialog에 `set_focus()` 호출
- 데이터 입력: `dlg.type_keys()` 메서드 사용
- 좌표 사용 안 함 → 해상도 독립적 ✅
- 안정적이고 간단한 구현 ✅

**핵심 교훈:**
- MFC 커스텀 컨트롤은 표준 메시지가 작동하지 않을 수 있음
- pywinauto의 고수준 메서드가 가장 안정적
- Dialog 구조를 이해하면 더 나은 접근 가능

---

**작성자:** Claude Code
**테스트 환경:** Windows 11, Python 3.12, pywinauto 0.6.8
**대상 애플리케이션:** 케이렙 365 - 사원등록
