# 부양가족정보 입력 테스트 결과 요약

## 성공한 방법

### ✅ **Attempt 43: dlg.type_keys() 방식** (성공!)

**방법:**
```python
# 1. 탭 선택
tab_auto.select_tab("기본사항")
tab_auto.select_tab("부양가족정보")

# 2. Down 키로 입력 행 선택 (SendMessage 사용)
win32api.SendMessage(spread_hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
win32api.SendMessage(spread_hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)

# 3. dlg.type_keys()로 데이터 입력
for value in ["4", "강부양", "내", "2019"]:
    dlg.type_keys(value, with_spaces=False, pause=0.05)
    dlg.type_keys("{TAB}", pause=0.05)

# 4. Enter로 확정
dlg.type_keys("{ENTER}", pause=0.05)
```

**결과:**
- 연말관계: 4
- 성명: 강부양
- 내/외국: 내
- 년도: 2019
- ✅ 모든 데이터가 성공적으로 입력되고 확정됨

**핵심 포인트:**
- `dlg.type_keys()` 사용 (다이얼로그 윈도우 전체에 키 입력)
- `spread.type_keys()`는 "ElementNotVisible" 오류 발생
- **관리자 권한으로 실행 필수** (UIPI 보안 때문)

---

## 실패한 방법들

### ❌ Attempt 39: 클립보드 복사-붙여넣기
- 클립보드에 복사 후 Ctrl+V
- SendMessage로 Ctrl+V 전송
- **실패**: 데이터 입력 안 됨

### ❌ Attempt 40: SendInput (전역 키보드)
- ctypes SendInput으로 유니코드 문자 전송
- **실패**: 데이터 입력 안 됨

### ❌ Attempt 41: SetForegroundWindow + SendInput
- 윈도우를 포그라운드로 활성화 후 SendInput
- **실패**: 데이터 입력 안 됨

### ❌ Attempt 42: spread.type_keys()
- fpUSpread80 컨트롤에 직접 type_keys
- **실패**: ElementNotVisible 오류

### ❌ Attempts 34-36: WM_CHAR + SendMessage
- fpUSpread80 HWND에 WM_CHAR 메시지 전송
- **실패**: 데이터 입력 안 됨

---

## 중요 발견사항

### 1. 관리자 권한 문제
**증상:**
- 일반 권한 터미널: 탭 선택 안 됨, 입력 안 됨
- 관리자 권한 터미널: 모든 기능 정상 작동

**원인:**
- 사원등록 프로그램이 관리자 권한으로 실행 중
- Windows UIPI(User Interface Privilege Isolation) 보안 메커니즘
- 낮은 권한 프로세스 → 높은 권한 프로세스 메시지 차단

**해결:**
- **관리자 권한 터미널에서 스크립트 실행 필수**

### 2. 탭 선택 메커니즘
**방법:**
```python
# TabAutomation 클래스 사용
tab_auto = TabAutomation()
tab_auto.connect()
tab_auto.select_tab("부양가족정보")
```

**내부 동작:**
- 탭 컨트롤(Afx:TabWnd:*) 찾기
- 탭 좌표에 WM_LBUTTONDOWN/UP 메시지 전송
- 클라이언트 좌표 사용 (50, 15), (150, 15) 등

**안정성:**
- 관리자 권한에서 매우 안정적
- 일반 권한에서는 작동하지 않음

### 3. 입력 방식 차이
**실패한 방식:**
- SendMessage(spread_hwnd, WM_CHAR, ...)
- SendInput (전역 키보드)
- Ctrl+V 클립보드 붙여넣기

**성공한 방식:**
- `dlg.type_keys()` - pywinauto 고수준 메서드
- 다이얼로그 윈도우 전체에 키 입력
- 내부적으로 올바른 포커스 처리

### 4. fpUSpread80 특성
**발견사항:**
- 직접 HWND에 메시지 전송은 작동하지 않음
- 탭 선택 후 포커스가 자동으로 스프레드로 이동
- 이 상태에서 dlg.type_keys()가 올바르게 작동

---

## 권장 사용 방법

### 최종 코드 템플릿

```python
from pywinauto import application
from tab_automation import TabAutomation
import win32api
import win32con
import time

# 애플리케이션 연결
app = application.Application(backend="win32")
app.connect(title="사원등록")
dlg = app.window(title="사원등록")

# 탭 자동화 초기화
tab_auto = TabAutomation()
tab_auto.connect()

# 1. 기본사항 탭 선택
tab_auto.select_tab("기본사항")
time.sleep(1.0)

# 2. 부양가족정보 탭 선택
tab_auto.select_tab("부양가족정보")
time.sleep(1.0)

# 3. 스프레드 컨트롤 찾기
spread = None
for ctrl in dlg.descendants():
    try:
        if ctrl.class_name() == "fpUSpread80":
            spread = ctrl
            break
    except:
        pass

# 4. Down 키로 입력 행 이동
win32api.SendMessage(spread.handle, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
win32api.SendMessage(spread.handle, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
time.sleep(0.5)

# 5. 데이터 입력
data = ["4", "홍길동", "내", "2020"]
for idx, value in enumerate(data):
    dlg.type_keys(value, with_spaces=False, pause=0.05)
    time.sleep(0.3)

    if idx < len(data) - 1:
        dlg.type_keys("{TAB}", pause=0.05)
        time.sleep(0.2)

# 6. Enter로 확정
dlg.type_keys("{ENTER}", pause=0.05)
time.sleep(0.5)
```

### 주의사항

1. **반드시 관리자 권한으로 실행**
   ```bash
   # 관리자 권한 터미널에서 실행
   uv run python your_script.py
   ```

2. **적절한 대기 시간**
   - 탭 전환 후: 1.0초
   - Down 키 후: 0.5초
   - 각 필드 입력 후: 0.3초
   - Tab 키 후: 0.2초

3. **연말관계 코드**
   - 0: 본인
   - 1: 부모
   - 2: 배우자부모
   - 3: 배우자
   - 4: 자녀

4. **에러 처리**
   - 탭 컨트롤을 찾지 못하면 실패
   - 스프레드 컨트롤을 찾지 못하면 실패
   - 관리자 권한 없으면 입력 실패

---

## 테스트 이력

- Attempt 08-18: SPR32DU80EditHScroll 대상 (실패)
- Attempt 11-15: fpUSpread80 발견, 좌표 클릭 (성공)
- Attempt 19-33: 키보드 전용 방식 시도 (모두 실패)
- Attempt 34-36: 연말관계 코드 + 탭 전환 (실패)
- Attempt 37-38: 탭 선택 진단 (관리자 권한 문제 발견)
- Attempt 39-42: 다양한 입력 방식 시도 (모두 실패)
- **Attempt 43: dlg.type_keys() (성공!)** ✅

---

## 결론

**성공 조건:**
1. ✅ 관리자 권한 실행
2. ✅ TabAutomation으로 탭 선택
3. ✅ Down 키로 행 이동
4. ✅ `dlg.type_keys()`로 입력
5. ✅ 적절한 대기 시간

**핵심 교훈:**
- fpUSpread80는 특수한 컨트롤로 일반적인 메시지 전송 방식 불가
- pywinauto의 고수준 메서드(`dlg.type_keys`)가 올바른 접근 방식
- 관리자 권한이 모든 기능의 전제 조건

---

생성 일시: 2025-10-30
최종 성공: Attempt 43
