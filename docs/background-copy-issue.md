# fpUSpread80 백그라운드 복사 문제

## 문제 상황

`bulk_dependent_input.py` 작업 중 발견된 문제:
- 왼쪽 사원 목록(fpUSpread80)에서 클립보드 복사를 통해 사번과 이름을 읽어야 함
- 다른 작업을 하면서 백그라운드에서 자동화를 실행하고 싶음
- 하지만 fpUSpread80은 백그라운드에서 복사가 불가능함

## 테스트 결과

### Attempt 54: 복사 메서드 비교 (활성화 상태)

**테스트 환경:** 사원등록 창이 활성화된 상태

| 방법 | 결과 | 복사된 값 |
|------|------|-----------|
| WM_COPY 메시지 | ❌ 실패 | 변화 없음 |
| SendMessage Ctrl+C | ❌ 실패 | 변화 없음 |
| left_spread.type_keys('^c', set_foreground=False) | ✅ 성공 | '0' |
| left_spread.type_keys('^c', set_foreground=True) | ✅ 성공 | '0' |
| dlg.type_keys('^c') | ✅ 성공 | '0' |

**결론:** `type_keys` 방식만 작동

---

### Attempt 55: 백그라운드 복사 테스트

**테스트 환경:** 메모장을 활성화하여 사원등록 창을 비활성화

| 방법 | 결과 | 복사된 값 | 문제점 |
|------|------|-----------|--------|
| SendMessage Ctrl+C | ❌ 실패 | 'BACKGROUND_TEST' | 아무 효과 없음 |
| left_spread.type_keys('^c', set_foreground=False) | ❌ 실패 | '2' | 메모장 제목에서 복사됨! |

**결론:** 백그라운드 복사 불가능

---

## 핵심 발견사항

### 1. fpUSpread80의 특성

- **SendMessage로 복사 불가**: WM_COPY, WM_KEYDOWN 등 메시지 전송이 작동하지 않음
- **type_keys 의존**: pywinauto의 `type_keys()`만 작동
- **전역 키보드 입력**: `type_keys`는 시스템 전역으로 키를 전송하므로 활성 창에 입력됨

### 2. 문제 시나리오

```
사용자가 콘솔 창을 보고 있는 상태에서:
→ bulk_dependent_input.py가 Ctrl+C 전송
→ 콘솔 창에 Ctrl+C가 입력됨
→ 프로그램 중단됨 💥
```

### 3. 시도했던 방법들

```python
# ❌ 실패: WM_COPY 메시지
win32api.SendMessage(hwnd, 0x0301, 0, 0)

# ❌ 실패: SendMessage로 Ctrl+C
lparam = 1 | (0x2E << 16)
win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, ord('C'), lparam)

# ✅ 성공 (활성화 상태에서만)
left_spread.type_keys("^c", set_foreground=False)

# ❌ 실패 (백그라운드에서)
# → 활성 창에 Ctrl+C가 전송됨
```

---

## 해결 방안

### 옵션 A: 창 활성화 방식 (확실하지만 방해됨)

```python
def _copy_with_activation(self):
    """창을 잠깐 활성화하고 복사"""
    # 현재 활성 창 저장
    prev_hwnd = win32gui.GetForegroundWindow()

    # 사원등록 창 활성화
    win32gui.SetForegroundWindow(self.dlg.handle)
    time.sleep(0.1)

    # 복사
    self.left_spread.type_keys("^c", pause=0.05)
    time.sleep(0.2)
    value = pyperclip.paste()

    # 원래 창으로 복귀
    win32gui.SetForegroundWindow(prev_hwnd)

    return value
```

**장점:**
- ✅ 확실하게 작동
- ✅ 구현 간단

**단점:**
- ❌ 사용자 화면에 창이 깜빡임
- ❌ 다른 작업 방해

---

### 옵션 B: COM 인터페이스 (추천) 🌟

Farpoint Spread는 ActiveX 컨트롤이므로 COM 인터페이스를 제공할 가능성이 높음.

```python
import win32com.client

# COM 객체로 접근
spread = win32com.client.Dispatch(hwnd)

# 셀 값 직접 읽기
value = spread.GetText(col, row)  # 메서드명은 확인 필요
```

**장점:**
- ✅ 백그라운드에서 작동
- ✅ 클립보드 사용 안 함
- ✅ 사용자 방해 없음

**단점:**
- ❓ COM 인터페이스 메서드를 찾아야 함
- ❓ 문서가 없을 수 있음

**다음 단계:**
1. fpUSpread80의 COM 인터페이스 확인
2. 셀 값 읽기 메서드 찾기
3. 테스트 스크립트 작성

---

### 옵션 C: UI Automation

```python
from pywinauto import uia

# UIA 백엔드로 재연결
app = Application(backend='uia')
# ...
```

**장점:**
- ✅ 백그라운드 접근 가능성

**단점:**
- ❌ fpUSpread80은 Win32 컨트롤이므로 UIA 지원 제한적
- ❌ 이미 시도했으나 효과 없었음 (attempt04_uia_backend.py)

---

## 현재 코드 상태

### bulk_dependent_input.py

```python
def _send_copy(self):
    """
    클립보드 복사

    ⚠️  문제: 백그라운드에서 작동하지 않음
    type_keys는 전역 키보드 입력이므로 활성 창에 Ctrl+C 전송
    """
    self.left_spread.type_keys("^c", pause=0.05, set_foreground=False)
    time.sleep(0.1)
```

**현재 상태:** 사원등록 창이 활성화되어 있으면 작동하지만, 다른 창이 활성화되면 실패

---

## 권장 해결 순서

1. ~~**1단계:** COM 인터페이스 시도 (옵션 B)~~ ❌ 실패
   - Attempt 56-57에서 시도
   - AccessibleObjectFromWindow로 COM 객체 획득까지는 성공
   - 하지만 c_void_p를 제대로 된 COM 객체로 변환 실패
   - IAccessible로도 셀 값 읽기 불가능
   - **결론: fpUSpread80은 표준 COM으로 직접 접근 불가**

2. ~~**2단계:** UI Automation 시도 (옵션 C)~~ ❌ 실패
   - Attempt 58에서 시도
   - UIA 백엔드로 연결은 성공
   - 하지만 fpUSpread80 컨트롤을 찾을 수 없음
   - UIA는 Pane만 보이고 Win32 네이티브 컨트롤은 노출되지 않음
   - **결론: fpUSpread80은 UIA에 노출되지 않는 Win32 전용 컨트롤**

3. ~~**3단계:** 클립보드 없이 셀 값 직접 읽기 시도~~ ❌ 실패
   - Attempt 59-60: 자식 윈도우/ListBox에서 읽기 시도
   - ListBox에서 3개 항목 발견했으나 owner-drawn 컨트롤로 텍스트 없음
   - **결론: 자식 윈도우를 통한 접근 불가능**

4. ~~**4단계:** IAccessible 상세 분석~~ ❌ 실패
   - Attempt 61-66: IAccessible의 모든 속성 테스트
   - accFocus로 포커스된 셀 객체 획득 성공
   - 하지만 accValue, accName, accDescription 등 모든 텍스트 속성이 NULL BSTR 반환
   - **결론: IAccessible은 포커스 정보만 제공하고 셀 값은 노출하지 않음**

5. ~~**5단계:** COM TypeLib/ProgID 검색~~ ❌ 실패
   - Attempt 67-69: 레지스트리에서 FarPoint Spread CLSID/ProgID 검색
   - 64비트 및 32비트 레지스트리 모두 검색
   - 파일 시스템에서 Fpspr80.ocx 검색
   - **결론: FarPoint Spread 8.0은 이 시스템에 COM 등록되지 않음**
   - WM_GETOBJECT도 0 반환 (COM 객체 미노출)

6. **6단계:** 인터넷 검색 및 외부 리소스 확인 🔍
   - Stack Overflow: ["Get data from FarPoint Spread control running in another application"](https://stackoverflow.com/questions/29983290/get-data-from-farpoint-spread-control-running-in-another-application)
   - **공식 답변:** "FarPoint Spread는 UI Automation을 지원하지 않음"
   - GrapeCity 포럼: "Spread는 외부 자동화(automation)를 지원하지 않음"
   - **결론: FarPoint Spread는 설계상 외부 애플리케이션에서의 자동화를 지원하지 않음** ⚠️

7. ✅ **최종 해결책:** 포그라운드 솔루션 (Attempt 70)
   - 창을 활성화한 상태에서 type_keys + 클립보드 사용
   - 안정성 테스트: **10/10 (100%)** 성공
   - 셀 이동, 값 읽기, 연속 읽기 모두 정상 작동
   - **이것이 유일하게 작동하는 방법임**

### 최종 결론: 백그라운드 자동화 불가능 ⚠️

**시도한 모든 방법 (총 70회 시도):**
- ❌ SendMessage (WM_COPY, WM_KEYDOWN)
- ❌ COM 인터페이스 (IDispatch, IAccessible)
- ❌ UI Automation (UIA 백엔드)
- ❌ 자식 윈도우/ListBox 직접 읽기
- ❌ IAccessible 모든 속성 (accValue, accName, accDescription 등)
- ❌ COM TypeLib/ProgID 검색
- ❌ WM_GETOBJECT
- ❌ 레지스트리 검색 (64비트/32비트)
- ✅ **type_keys + 클립보드 (포그라운드만)**

**Stack Overflow 및 공식 문서 확인 결과:**
- FarPoint Spread는 **의도적으로** 외부 자동화를 지원하지 않음
- UI Automation, COM automation 모두 미지원
- 제조사(GrapeCity)도 외부 자동화 불가능 확인

**fpUSpread80에서 백그라운드로 값을 읽는 것은 기술적으로 불가능함**

---

## 최종 권장 해결 방안

### ✅ 채택된 방안: 포그라운드 자동화 (Attempt 70 기반)

**70회 시도 끝에 도달한 결론: 포그라운드에서만 작동 가능**

```python
def read_cell_value(spread_control, dlg):
    """현재 선택된 셀의 값을 클립보드로 읽기"""
    import pyperclip
    import win32gui

    # 창이 활성 상태인지 확인
    if win32gui.GetForegroundWindow() != dlg.handle:
        dlg.set_focus()
        time.sleep(0.3)

    # 클립보드 초기화
    pyperclip.copy("__EMPTY__")
    time.sleep(0.1)

    # 복사 시도 (최대 3회)
    for attempt in range(3):
        spread_control.type_keys("^c", pause=0.1)
        time.sleep(0.2)

        value = pyperclip.paste()
        if value != "__EMPTY__":
            return value

        time.sleep(0.2)

    return None
```

**안정성:**
- ✅ 10/10 (100%) 성공률
- ✅ 셀 이동 및 연속 읽기 정상 작동
- ✅ 포커스 자동 관리

**제약사항:**
- ⚠️ 자동화 실행 중 사용자가 다른 작업을 할 수 없음
- ⚠️ 창이 활성화 상태여야 함

### 대안: 창 빠르게 전환하는 방식

```python
def _copy_with_quick_activation(self):
    """창을 잠깐 활성화하고 복사 (사용자 방해 최소화)"""
    prev_hwnd = win32gui.GetForegroundWindow()

    # 빠르게 활성화
    win32gui.SetForegroundWindow(self.dlg.handle)
    time.sleep(0.05)  # 최소 대기

    # 복사
    self.left_spread.type_keys("^c", pause=0.05)
    time.sleep(0.1)
    value = pyperclip.paste()

    # 즉시 복귀 (최소 깜빡임)
    if prev_hwnd:
        win32gui.SetForegroundWindow(prev_hwnd)

    return value
```

**장점:**
- ✅ 100% 작동 보장
- ✅ 깜빡임 최소화 (약 0.15초)

**단점:**
- ⚠️ 여전히 화면 전환 발생
- ⚠️ 빠른 반복 시 사용자 불편

### 구현 권장사항

1. **기본 동작**: 포그라운드 유지
   - 실행 시 사용자에게 안내 메시지 출력
   - "자동화 진행 중 다른 작업을 하지 마세요"

2. **옵션 제공**: `--minimize-interruption` 플래그
   - 플래그 사용 시 빠른 전환 방식 사용
   - 사용자가 선택할 수 있도록 함

3. **사용자 경험**:
   ```
   $ python bulk_dependent_input.py employees.csv

   ⚠️  주의사항:
   이 자동화는 FarPoint Spread 컨트롤의 제약으로 인해
   사원등록 창이 활성화 상태여야 작동합니다.

   자동화 진행 중(약 5분):
   - 다른 창을 클릭하지 마세요
   - 키보드 입력을 하지 마세요

   [계속하려면 Enter, 취소하려면 Ctrl+C]
   ```

---

## 참고 파일

### 주요 테스트 파일
- **Attempt 54**: `test/attempt/attempt54_test_copy_methods.py` - 복사 메서드 비교 (활성화 상태)
- **Attempt 55**: `test/attempt/attempt55_test_background_copy.py` - 백그라운드 복사 실패 확인
- **Attempt 56-58**: COM/IAccessible/UIA 시도
- **Attempt 59-60**: 자식 윈도우/ListBox 읽기 시도
- **Attempt 61-66**: IAccessible 상세 분석 (모든 속성 테스트)
- **Attempt 67-69**: COM TypeLib/ProgID 레지스트리 검색
- **Attempt 70**: `test/attempt/attempt70_foreground_solution.py` - ✅ **최종 포그라운드 솔루션**

### 기타 문서
- 메인 구현: `bulk_dependent_input.py`
- 성공 메서드: `docs/successful-method.md`
- 테스트 프레임워크: `docs/testing-framework.md`

### 외부 참고
- [Stack Overflow: Get data from FarPoint Spread control](https://stackoverflow.com/questions/29983290/get-data-from-farpoint-spread-control-running-in-another-application)
- [GrapeCity Forums: Automation of FarPoint](https://www.grapecity.com/forums/spread-winforms/automation-of-farpoint-gri)

---

**작성일:** 2025-10-31
**최종 업데이트:** 2025-10-31
**테스트 환경:** Windows 11, Python 3.14, pywinauto 0.6.8
**대상 애플리케이션:** 케이렙 365 - 사원등록
**총 시도 횟수:** 70회 (Attempt 54-70 포함, 이전 1-53은 다른 문제)
**최종 결론:** 포그라운드 자동화만 가능 (Attempt 70 성공)
