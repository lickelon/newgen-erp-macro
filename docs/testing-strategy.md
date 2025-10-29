# 연말정산 자동화 테스트 전략

## 목표

**연말정산추가자료입력** 프로그램에서 **부양가족 탭**을 자동으로 선택

## 제약사항

### 1. 마우스 직접 이동 금지
- ❌ `pyautogui.click()` - 물리적 마우스 이동
- ❌ `mouse.click()` - 물리적 마우스 이동
- ❌ 절대 좌표 물리적 클릭
- ✅ `pywinauto`의 `click_input()` - 윈도우 메시지 사용
- ✅ `win32api.SendMessage()` - 윈도우 메시지 직접 전송

### 2. pywinauto 사용
- MFC 기반 애플리케이션이므로 win32 백엔드 사용
- 32비트 애플리케이션이지만 64비트 Python으로 제어 가능 (경고 무시)

### 3. 윈도우 메시지 방식
- 마우스 커서를 움직이지 않고 윈도우 핸들에 직접 메시지 전송

## 테스트 방법

### 1. 스크린샷 기반 평가
- 매 단위 실행마다 **스크린샷 촬영**
- Claude가 **직접 이미지 확인**하여 부양가족 탭 선택 여부 평가
- 성공할 때까지 **계속 시도**

### 2. 반복적 시도
- 다양한 접근 방법을 순차적으로 시도
- 각 시도는 독립적인 스크립트로 관리
- 실패 시 다음 방법으로 진행

## 파일 구조

```
newgen-erp-macro/
├── docs/
│   └── testing-strategy.md          # 이 문서
├── test/
│   ├── __init__.py
│   ├── attempt/
│   │   ├── __init__.py
│   │   ├── attempt01_click_children.py
│   │   ├── attempt02_send_message.py
│   │   ├── attempt03_coordinate_scan.py
│   │   └── ...
│   ├── image/                        # 스크린샷 저장
│   │   ├── attempt01_00_initial.png
│   │   ├── attempt01_01_button.png
│   │   └── ...
│   └── capture.py                    # 캡처 유틸리티
├── test.py                           # 메인 실행 스크립트
└── main.py                           # 최종 자동화 스크립트
```

## Attempt 스크립트 구조

### 파일명 규칙
`attempt{번호}_{요약}.py`

예시:
- `attempt01_click_children.py` - 자식 요소 클릭
- `attempt02_send_message.py` - WM_LBUTTONDOWN 메시지
- `attempt03_coordinate_scan.py` - 좌표 스캔

### 함수 시그니처

```python
"""
시도 N: 방법 설명
"""
import sys
import time

def run(dlg, capture_func):
    """
    Args:
        dlg: pywinauto 윈도우 객체 (연말정산추가자료입력)
        capture_func: 스크린샷 함수
                     signature: (filename: str) -> None
                     이미지는 test/image/에 자동 저장

    Returns:
        dict: {
            "success": bool,      # 시도 성공 여부
            "message": str        # 결과 메시지
        }
    """
    print("\n" + "="*60)
    print("시도 N: 방법 설명")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attemptN_00_initial.png")

        # 테스트 로직
        # ...

        # 각 단계마다 스크린샷
        capture_func("attemptN_01_step1.png")

        return {"success": True, "message": "완료"}

    except Exception as e:
        return {"success": False, "message": f"오류: {e}"}
```

## test.py 구조

```python
from pywinauto import application
from test.capture import capture_window

# 연결
app = application.Application(backend="win32")
app.connect(title="연말정산추가자료입력")
dlg = app.window(title="연말정산추가자료입력")
hwnd = dlg.handle

# capture 함수 생성
def capture_func(filename):
    capture_window(hwnd, filename)

# 시도 1 실행
from test.attempt.attempt01_click_children import run as attempt01
result = attempt01(dlg, capture_func)
print(f"결과: {result}")

# 실패 시 시도 2 실행
if not result["success"]:
    from test.attempt.attempt02_send_message import run as attempt02
    result = attempt02(dlg, capture_func)
    print(f"결과: {result}")
```

## 시도한 방법들

### ❌ 실패한 방법들

1. **TCM_SETCURSEL 메시지** (0x130C)
   - 표준 탭 컨트롤 메시지
   - MFC 커스텀 탭에서 작동 안 함

2. **WM_LBUTTONDOWN/UP 메시지**
   - 일반 클릭 메시지
   - 탭 선택 이벤트 발생 안 함

3. **키보드 입력** (VK_RIGHT)
   - 화살표 키로 탭 이동
   - 포커스 문제로 작동 안 함

4. **물리적 마우스 클릭**
   - 좌표 기반 클릭
   - 제약사항 위반

5. **부양가족탭불러오기 버튼**
   - 화면 밖에 위치 (ElementNotVisible)
   - 클릭 불가

### 🔄 진행 중

- pywinauto의 다양한 클릭 메서드 테스트
- 탭 컨트롤 자식 요소 분석 및 클릭

## 탭 컨트롤 정보

### 기본 정보
- **클래스명**: `Afx:TabWnd:330000:8:10003:10`
- **위치**: 가변 (예: L=3845, T=1183, R=5713, B=1215)
- **높이**: 32px
- **타입**: MFC 커스텀 탭 컨트롤

### 탭 목록 (왼쪽부터)
1. 소득정보 (현재 선택됨)
2. **부양가족** ← 목표
3. 신용카드 등
4. 의료비
5. 기부금
6. ...

## 디버깅 팁

### 스크린샷 확인
```python
from PIL import Image
img = Image.open("test/image/attempt01_01_button.png")
img.show()
```

### 컨트롤 정보 출력
```python
dlg.print_control_identifiers()
```

### 탭 컨트롤 자식 탐색
```python
tab = dlg.child_window(class_name="Afx:TabWnd:330000:8:10003:10")
for child in tab.descendants():
    print(child.class_name(), child.rectangle())
```

## 다음 시도 계획

1. 탭 컨트롤의 모든 자식 요소 클릭
2. 다양한 pywinauto 클릭 메서드 테스트
3. 탭 영역 좌표 스캔 (메시지 방식)
4. 부모 윈도우에 직접 메시지 전송
5. UI Automation (UIA 백엔드) 시도
