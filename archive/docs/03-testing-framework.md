# 테스트 프레임워크

스크린샷 기반 테스트 프레임워크와 attempt 스크립트 구조

## 파일 구조

```
newgen-erp-macro/
├── docs/
│   ├── overview.md                   # 프로젝트 개요
│   ├── window-architecture.md        # 윈도우 구조
│   ├── tab-automation.md             # 탭 자동화
│   ├── testing-framework.md          # 테스트 프레임워크 (이 문서)
│   └── development-guide.md          # 개발 가이드
├── test/
│   ├── __init__.py
│   ├── attempt/
│   │   ├── __init__.py
│   │   ├── attempt01_click_children.py
│   │   ├── attempt02_find_tab_control.py
│   │   ├── attempt03_sendmessage_tab.py
│   │   ├── attempt04_uia_backend.py
│   │   ├── attempt05_keyboard_input.py
│   │   ├── attempt06_direct_tab_hwnd.py       # ✅ 성공
│   │   └── attempt07_robust_tab_find.py
│   ├── image/                        # 스크린샷 저장
│   │   ├── attempt01_00_initial.png
│   │   ├── attempt06_01_tab_2_부양가족정보.png
│   │   └── ...
│   └── capture.py                    # 캡처 유틸리티
├── test.py                           # 메인 테스트 스크립트
├── tab_automation.py                 # 탭 자동화 모듈
└── main.py                           # 최종 자동화 스크립트
```

## Attempt 스크립트 구조

### 파일명 규칙
`attempt{번호}_{요약}.py`

예시:
- `attempt01_click_children.py` - 자식 요소 클릭
- `attempt02_find_tab_control.py` - 탭 컨트롤 찾기
- `attempt06_direct_tab_hwnd.py` - 탭 컨트롤에 직접 메시지 전송 (성공)

### 함수 시그니처

```python
"""
시도 N: 방법 설명
"""
import time

def run(dlg, capture_func):
    """
    Args:
        dlg: pywinauto 윈도우 객체 (사원등록)
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
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
```

### 스크린샷 파일명 규칙

- `attemptN_00_initial.png` - 초기 상태
- `attemptN_01_step1.png` - 첫 번째 단계
- `attemptN_02_step2.png` - 두 번째 단계
- ...

**네이밍 팁**:
- 숫자 앞에 0 패딩 (정렬을 위해)
- 의미있는 단계 이름 사용
- 예: `attempt06_01_tab_2_부양가족정보.png`

## test.py 구조

```python
"""
사원등록 자동화 테스트 메인 스크립트
"""
from pywinauto import application
import sys
from test.capture import capture_window

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("="*70)
    print("사원등록 자동화 테스트")
    print("="*70)

    # 연결
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        hwnd = dlg.handle

        print(f"\n✓ 사원등록 윈도우 연결 성공 (HWND: {hwnd})\n")

    except Exception as e:
        print(f"\n✗ 연결 실패: {e}")
        return

    # capture 함수 생성
    def capture_func(filename):
        capture_window(hwnd, filename)

    # 시도 1 실행
    from test.attempt.attempt01_click_children import run as attempt01
    result = attempt01(dlg, capture_func)
    print(f"결과: {result}")

    # 실패 시 시도 2 실행
    if not result["success"]:
        from test.attempt.attempt02_find_tab_control import run as attempt02
        result = attempt02(dlg, capture_func)
        print(f"결과: {result}")

    # ... 계속

if __name__ == "__main__":
    main()
```

## 캡처 유틸리티 (test/capture.py)

```python
"""
스크린샷 캡처 유틸리티
"""
import os
from PIL import ImageGrab
import win32gui

def capture_window(hwnd, filename):
    """
    특정 윈도우의 스크린샷 캡처

    Args:
        hwnd: 윈도우 핸들
        filename: 파일명 (경로 제외)
    """
    # 저장 디렉토리
    image_dir = os.path.join(os.path.dirname(__file__), "image")
    os.makedirs(image_dir, exist_ok=True)

    # 윈도우 영역 가져오기
    rect = win32gui.GetWindowRect(hwnd)

    # 스크린샷 촬영
    img = ImageGrab.grab(bbox=rect)

    # 저장
    filepath = os.path.join(image_dir, filename)
    img.save(filepath)

    return filepath
```

## 스크린샷 기반 평가

### 프로세스

```
1. attempt 스크립트 실행
   ↓
2. 각 단계마다 스크린샷 촬영
   ↓
3. test/image/ 폴더에 저장
   ↓
4. Claude가 이미지 확인
   ↓
5. 탭 선택 여부 판단
   ↓
6. 성공 시 종료, 실패 시 다음 attempt
```

### 평가 기준

**성공**:
- 목표 탭이 선택됨 (시각적으로 확인)
- 마우스 커서가 움직이지 않음 (사용자 확인)

**실패**:
- 탭이 변경되지 않음
- 에러 발생
- 마우스 커서가 움직임 (제약사항 위반)

## 새 Attempt 추가하기

### 1. 파일 생성

```bash
touch test/attempt/attemptN_description.py
```

### 2. 템플릿 복사

```python
"""
시도 N: [방법 설명]
"""
import time

def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 N: [방법 설명]")
    print("="*60)

    try:
        capture_func("attemptN_00_initial.png")

        # 여기에 테스트 로직 작성

        return {"success": True, "message": "완료"}
    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
```

### 3. test.py에 추가

```python
# 시도 N: [방법 설명]
from test.attempt.attemptN_description import run as attemptN
result = attemptN(dlg, capture_func)
print(f"결과: {result}")
```

### 4. 실행 및 평가

```bash
uv run python test.py
```

스크린샷 확인:
- `test/image/attemptN_*.png`

## 디버깅

### 스크린샷 수동 확인

```python
from PIL import Image
img = Image.open("test/image/attempt06_01_tab_2_부양가족정보.png")
img.show()
```

### 컨트롤 정보 출력

```python
dlg.print_control_identifiers()
```

### 탭 컨트롤 자식 탐색

```python
tab = dlg.child_window(class_name="Afx:TabWnd:...")
for child in tab.descendants():
    print(child.class_name(), child.rectangle())
```

## 성공 사례: attempt06

```
시도 6 실행
→ 탭 컨트롤 HWND 찾기
→ WM_LBUTTONDOWN 전송 (x=150, y=15)
→ WM_LBUTTONUP 전송
→ 스크린샷 촬영
→ 결과: 부양가족정보 탭 선택됨 ✅
```

**스크린샷**:
- `attempt06_00_initial.png` - 부양가족정보 탭
- `attempt06_01_tab_2_부양가족정보.png` - 부양가족정보 탭 (변화 없음)
- `attempt06_01_tab_3_소득자료.png` - 소득자료 탭 선택됨 ✅

## 참고

- 모든 attempt 스크립트: `test/attempt/`
- 모든 스크린샷: `test/image/`
- 캡처 유틸리티: `test/capture.py`
