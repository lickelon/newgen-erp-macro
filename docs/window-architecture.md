# 윈도우 아키텍처

사원등록 프로그램의 윈도우 계층 구조와 컨트롤 정보

## 윈도우 계층 구조 (Spy++ 분석)

```
Window 000A0828 "사원등록" Afx:RibbonBarcd0000:8:10003:10
├─ Window 000608EE "" Afx:TabWnd:cd0000:8:10003:10  ← 탭 컨트롤
│  ├─ Window 000608E6 "" Button
│  ├─ Window 000608DA "" Button
│  └─ Window 000608CE "" Button
├─ Window 00060A40 " 기본사항 " #32770 (Dialog)
├─ Window 002306B4 " 부양가족정보 " #32770 (Dialog)
│  └─ Window 000609BA "" GwProPageClass
│     └─ Window 000609BE "" AfxWnd90u
│        └─ Window 00060ADC "" fpUSpread80  ← 부양가족 입력 표
└─ Window 00060D8 "" Afx:FrameOrView90u
```

## 주요 컨트롤

### 1. 메인 윈도우
- **클래스명**: `Afx:RibbonBarcd0000:8:10003:10` (부분 가변)
- **제목**: "사원등록"
- **HWND**: 프로그램 실행마다 다름

### 2. 탭 컨트롤
- **클래스명 패턴**: `Afx:TabWnd:*` (부분 매칭 필요)
  - 예시: `Afx:TabWnd:cd0000:8:10003:10`
  - ⚠️ `cd0000` 부분은 프로그램 재시작 시 변경될 수 있음
- **HWND**: 프로그램 실행마다 다름 (예: 0x000608EE)
- **위치**: L=4255, T=1188, R=5713, B=1780 (가변)
- **크기**: W=1458, H=592
- **타입**: MFC 커스텀 탭 컨트롤
- **자식 요소**: 3개의 Button (탭 버튼)

#### 탭 목록
1. **기본사항** (x=50, y=15)
2. **부양가족정보** (x=150, y=15)
3. **소득자료** (x=250, y=15)

### 3. 부양가족 입력 표 (fpUSpread80)
- **클래스명**: `fpUSpread80`
- **HWND**: 프로그램 실행마다 다름 (예: 0x00060ADC)
- **타입**: Farpoint Spread 컨트롤 (엑셀 같은 스프레드시트)
- **위치**: "부양가족정보" 탭 내부
- **부모**: GwProPageClass → AfxWnd90u → fpUSpread80

#### 용도
- 부양가족 명세 데이터 입력
- 행/열 구조의 표 형식
- 주민등록번호, 성명, 관계 등 입력 필드

#### 데이터 입력 방법 (TODO)
- Farpoint Spread API 연구 필요
- SendMessage 또는 전용 메서드 사용
- 셀 선택 및 데이터 입력 자동화

## 안정적인 컨트롤 찾기

### 탭 컨트롤 찾기

HWND와 클래스명의 일부가 프로그램 재시작 시 바뀌므로 **부분 매칭** 사용:

```python
# 방법 1: 부분 클래스명 매칭 (권장)
tab_control = None
for ctrl in dlg.descendants():
    if ctrl.class_name().startswith("Afx:TabWnd:"):
        tab_control = ctrl
        break

# 방법 2: 직계 자식에서 찾기 (더 빠름)
for child in dlg.children():
    if "TabWnd" in child.class_name():
        tab_control = child
        break
```

### fpUSpread80 찾기

```python
# 클래스명으로 찾기
spread_controls = []
for ctrl in dlg.descendants():
    if ctrl.class_name() == "fpUSpread80":
        spread_controls.append(ctrl)

# 부양가족정보 탭 내부의 것 선택
# (총 3개 발견됨 - 각 탭에 하나씩)
```

## Spy++ 로그 분석 결과

### WM_NOTIFY 메시지
```
<000474> 000A0828 S WM_NOTIFY idCtrl: 1 pnmh: 012FF480
<000475> 000A0828 R WM_NOTIFY
```

- **idCtrl: 1** - 탭 컨트롤의 ID
- 탭 변경 후 부모 윈도우에 알림 메시지 전송
- 이 메시지는 탭 변경의 **결과**이지 원인이 아님

### 핵심 발견
- 탭 컨트롤 자체가 WM_LBUTTONDOWN/UP을 받으면 탭 변경
- 부모 윈도우가 아닌 **탭 컨트롤 HWND에 직접 전송**해야 함
- 좌표는 **클라이언트 좌표** (탭 컨트롤 내부 기준)

## 직계 자식 요소 목록

메인 윈도우의 직계 자식 151개 중 주요 항목:

- [0] Afx:RibbonStatusBar:cd0000:8:10003:10
- [1] Afx:RibbonBar:cd0000:8:10003:10
- [53] Afx:ToolBar:cd0000:8:10003:10
- [54] AfxFrameOrView90u
- **[55] Afx:TabWnd:cd0000:8:10003:10** ← 탭 컨트롤
- [59] #32770 - 기본사항 다이얼로그
- [60] GwProPageClass
- **[85] fpUSpread80** - 첫 번째 Spread 컨트롤
- [89] #32770 - 부양가족정보 다이얼로그
- **[93] fpUSpread80** - 두 번째 Spread 컨트롤 (부양가족용)
- [107] #32770 - 소득자료 다이얼로그
- **[145] fpUSpread80** - 세 번째 Spread 컨트롤

## 주의사항

1. **HWND는 매번 바뀜**: 하드코딩 금지
2. **클래스명 일부 가변**: 부분 매칭 필수
3. **좌표도 가변**: 절대 좌표 사용 금지, 클라이언트 좌표 사용
4. **여러 fpUSpread80 존재**: 올바른 것 선택 필요 (부양가족정보 탭 내부)

## 참고

- Spy++ 로그: `log.txt`
- Spy++ 스크린샷: `스크린샷 2025-10-30 113531.png`
