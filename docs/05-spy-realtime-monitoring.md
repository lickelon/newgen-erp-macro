# Spy++ 실시간 모니터링 가이드

자동화 스크립트 실행 중 윈도우 메시지를 실시간으로 모니터링하는 방법

## 설정 방법

### 1. Spy++ 메시지 로그 시작

1. **Spy++ 실행** (Visual Studio 포함)
2. **Spy → Messages** 메뉴 선택 (또는 Ctrl+M)
3. **파인더 도구**로 타겟 윈도우 선택
   - 사원등록 메인 윈도우
   - 또는 탭 컨트롤만 선택 (더 깔끔)
   - 또는 fpUSpread80 컨트롤 선택

### 2. 메시지 필터 설정

**필터 추천 설정:**

**탭 자동화 모니터링:**
```
✅ WM_LBUTTONDOWN (0x0201)
✅ WM_LBUTTONUP (0x0202)
✅ WM_NOTIFY (0x004E)
✅ TCM_* (탭 컨트롤 메시지들)
□ WM_PAINT (너무 많음 - 제외)
□ WM_TIMER (너무 많음 - 제외)
```

**fpUSpread80 데이터 입력 모니터링:**
```
✅ WM_SETTEXT (0x000C)
✅ WM_CHAR (0x0102)
✅ WM_KEYDOWN (0x0100)
✅ WM_KEYUP (0x0101)
✅ EM_* (Edit 컨트롤 메시지)
✅ 모든 메시지 (처음엔 전부 보는 것도 좋음)
```

### 3. 로그 시작

1. **OK** 버튼 클릭
2. Spy++ 메시지 뷰 창이 열림
3. **준비 완료** - 이제 스크립트 실행

## 사용 시나리오

### 시나리오 1: 탭 자동화 검증

**목적**: SendMessage가 제대로 전달되는지 확인

**절차**:
```bash
# 터미널 1: Spy++ 메시지 로그 시작 (탭 컨트롤 선택)

# 터미널 2: 스크립트 실행
uv run python tab_automation.py
```

**Spy++에서 확인할 내용:**
```
S 000608EE P WM_LBUTTONDOWN wParam:00000001 lParam:000F0096
R 000608EE
S 000608EE P WM_LBUTTONUP wParam:00000000 lParam:000F0096
R 000608EE
S 000A0828 P WM_NOTIFY idCtrl:1 ...
R 000A0828
```

- `000608EE`: 탭 컨트롤 HWND
- `lParam:000F0096`: LOWORD=150(x), HIWORD=15(y) → 부양가족정보 탭
- `WM_NOTIFY`: 탭 변경 후 부모에게 알림

### 시나리오 2: fpUSpread80 조사

**목적**: 부양가족 데이터 입력 시 어떤 메시지가 발생하는지 파악

**절차**:
```bash
# 터미널 1: Spy++ 메시지 로그 시작 (fpUSpread80 선택)
#           모든 메시지 필터링

# 터미널 2: 사원등록 프로그램 실행 후 대기

# 수동 작업: 마우스로 직접 셀 클릭 및 데이터 입력
# - 첫 번째 셀 클릭
# - "홍길동" 입력
# - 다음 셀로 이동

# Spy++에서 메시지 시퀀스 확인 및 저장
```

**예상 메시지:**
```
WM_LBUTTONDOWN → 셀 선택
WM_CHAR → 문자 입력
WM_SETTEXT → 텍스트 설정
또는 커스텀 메시지 (Farpoint 전용)
```

### 시나리오 3: 메시지 비교

**목적**: 수동 클릭 vs SendMessage 차이 확인

**A. 수동 클릭:**
```
1. Spy++ 로그 시작
2. 마우스로 직접 탭 클릭
3. 로그 저장: manual_click.txt
```

**B. SendMessage:**
```
1. Spy++ 로그 시작
2. uv run python tab_automation.py 실행
3. 로그 저장: sendmessage.txt
```

**비교:**
```bash
# 두 파일 비교
diff manual_click.txt sendmessage.txt
```

동일하면 → SendMessage가 정확히 작동
차이점 있으면 → 추가 메시지 필요

## 실전 팁

### 1. 타겟 윈도우 재선택

프로그램 재시작 시 HWND가 바뀌므로:
1. Spy++ 메시지 뷰에서 **Ctrl+M**
2. 새 HWND 선택
3. 로그 재시작

### 2. 로그 저장

```
Spy++ → File → Save
→ log_yyyymmdd_hhmmss.txt
```

나중에 분석 가능

### 3. 필터 프리셋

**프리셋 1: 최소 (깔끔)**
```
WM_LBUTTONDOWN
WM_LBUTTONUP
WM_NOTIFY
```

**프리셋 2: 디버깅 (상세)**
```
모든 메시지
단, WM_PAINT, WM_TIMER 제외
```

### 4. 좌표 확인

LPARAM에서 좌표 추출:
```
lParam: 0x000F0096
→ LOWORD = 0x0096 = 150 (x)
→ HIWORD = 0x000F = 15 (y)
```

Python으로 확인:
```python
lparam = 0x000F0096
x = lparam & 0xFFFF        # 150
y = (lparam >> 16) & 0xFFFF  # 15
print(f"x={x}, y={y}")
```

## 디버깅 시나리오

### 문제: 탭이 선택되지 않음

**진단:**
```
1. Spy++로 탭 컨트롤 모니터링
2. 스크립트 실행
3. WM_LBUTTONDOWN 메시지가 보이는가?
```

**경우 1: 메시지가 안 보임**
→ 잘못된 HWND로 전송됨
→ 탭 컨트롤 찾기 로직 수정

**경우 2: 메시지는 보이지만 탭이 안 바뀜**
→ 잘못된 좌표
→ lparam 값 확인 및 수정

**경우 3: 메시지 순서 문제**
→ DOWN 후 즉시 UP (대기 시간 부족)
→ time.sleep() 추가

### 문제: fpUSpread80 데이터 입력 실패

**진단:**
```
1. Spy++로 fpUSpread80 모니터링
2. 수동으로 셀 클릭 + 데이터 입력
3. 어떤 메시지 시퀀스가 발생하는지 관찰
4. 해당 메시지를 SendMessage로 재현
```

## 고급: 메시지 자동 분석

자동화 스크립트와 Spy++ 로그를 함께 분석:

```python
# spy_log_parser.py
def parse_spy_log(filename):
    """
    Spy++ 로그 파일 파싱
    """
    with open(filename, 'r', encoding='utf-16') as f:
        for line in f:
            if 'WM_LBUTTONDOWN' in line:
                # lparam 추출
                # 좌표 계산
                # 출력
                pass

# 사용
parse_spy_log('log.txt')
```

## 테스트 스크립트에 통합

```python
# test_with_spy_monitoring.py
import subprocess
import time

def test_tab_automation_with_spy():
    print("1. Spy++를 실행하고 탭 컨트롤을 모니터링하세요")
    print("2. 메시지 로그를 시작하세요 (Ctrl+M)")
    print("3. 준비되면 Enter를 누르세요...")
    input()

    print("\n자동화 스크립트 실행 중...")

    # 탭 자동화 실행
    from tab_automation import TabAutomation
    tab_auto = TabAutomation()
    tab_auto.connect()
    tab_auto.select_tab("부양가족정보")

    print("\nSpy++ 로그를 확인하세요:")
    print("- WM_LBUTTONDOWN 메시지가 보이나요?")
    print("- lparam 좌표가 (150, 15)인가요?")
    print("- WM_NOTIFY 메시지가 뒤따라 오나요?")

    input("\n확인 완료 후 Enter...")

if __name__ == "__main__":
    test_with_spy_monitoring()
```

## 참고

- Spy++ 매뉴얼: Visual Studio 문서
- 메시지 코드: [Windows Message Reference](https://docs.microsoft.com/en-us/windows/win32/winmsg/about-messages-and-message-queues)
- LPARAM 포맷: `MAKELONG(x, y)` = `(y << 16) | x`
