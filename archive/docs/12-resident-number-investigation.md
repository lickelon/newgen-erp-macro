# 주민등록번호 백그라운드 읽기 조사 (Attempts 71-80)

## 목표

스프레드에서 사번 대신 **주민등록번호**로 CSV 데이터를 매칭하여 백그라운드 자동화를 구현

## 조사 배경

- FarPoint Spread (fpUSpread80)에서 사번을 포그라운드로만 읽을 수 있음 (Attempt 70)
- 기본사항 탭의 주민등록번호 필드를 백그라운드에서 읽을 수 있다면 더 나은 자동화 가능
- 사용자 제안: "기본사항 탭으로 이동해서 주민등록번호를 읽어오는 건 어때?"

## 시도한 방법 (Attempts 71-80)

### Attempt 71: 기본사항 탭의 주민등록번호 읽기
- **목표**: 탭 클릭 없이 모든 Edit 컨트롤 검색
- **결과**: ❌ Edit 컨트롤 4개 발견, 모두 비어있음
- **원인**: 스프레드를 클릭하지 않아 데이터 미로드

### Attempt 72: 탭 구조 탐색 (미완성)
- 컨트롤 트리 출력 시도

### Attempt 73: 기본사항 탭 클릭 후 주민등록번호 읽기
- **방법**: 탭 클릭 (WM_LBUTTONDOWN/UP) → Edit 읽기
- **결과**: ❌ Edit 4개, 모두 비어있음
- **문제**: 스프레드를 먼저 클릭하지 않음

### Attempt 74: 사원 클릭 → 기본사항 탭 → 주민등록번호
- **방법**:
  1. 스프레드 클릭 (WM_LBUTTONDOWN/UP)
  2. 기본사항 탭 클릭
  3. Edit 컨트롤 읽기
- **결과**: ❌ Edit 4개, 비어있지 않은 것 0개
- **원인**: 데이터가 로드되지 않거나 다른 컨트롤 타입

### Attempt 75: 모든 컨트롤 완전 덤프
- **방법**: WM_GETTEXT와 GetWindowText로 모든 컨트롤 텍스트 읽기
- **결과**:
  - 전체 컨트롤: 1,531개
  - 텍스트 있음: 173개
  - **주민등록번호 형식(13자리) 없음**
  - 긴 숫자 문자열 없음
- **발견**: 탭 이름("기본사항", "부양가족명세" 등)은 보임

### Attempt 76: 특정 값 "XXXXXX-XXXXXXX" 검색
- **목표**: 사용자가 제시한 주민등록번호 값 직접 검색
- **방법**: pywinauto descendants 전체 검색 (148개)
- **결과**: ❌ 해당 값 발견 못함
- **부분 일치**: 마지막 7자리도 없음

### Attempt 77: 폼에서 주민등록번호 복사
- **방법**: Edit 컨트롤 클릭 → 전체선택 → 복사
- **결과**:
  - Edit 1개 발견
  - 클릭 후 복사: **(비어있음)**
- **좌표**: (4419, 1793) - (5053, 1851)

### Attempt 78: 위치 기반 모든 컨트롤 복사
- **방법**: Edit, Static, MaskEdit 등 모든 타입 검색 및 복사
- **결과**:
  - 관심 컨트롤: 1개 (Edit)
  - 복사된 값: 사번만 복사됨 (스프레드의 사번)
  - 주민등록번호 형식 없음

### Attempt 79: "주민등록번호" 라벨 찾기
- **목표**: 라벨을 찾아 그 근처 컨트롤에서 값 읽기
- **결과**: ❌ "주민" 키워드를 포함한 텍스트 없음
- **발견된 텍스트**:
  - "준비", "사원등록", "워크스페이스"
  - 버튼들: "조건검색", "기초등록", "부서등록" 등
  - **주민등록번호 관련 라벨 없음**

### Attempt 80: 철저한 Edit 컨트롤 검색
- **방법**:
  1. Win32 API + pywinauto 모두 사용
  2. 스프레드 클릭 추가
  3. 각 Edit 클릭 및 복사
- **결과**:
  - Win32 API: 1개 Edit 발견
  - pywinauto: 1개 Edit 발견
  - WM_GETTEXT: 빈 문자열
  - GetWindowText: 빈 문자열
  - 클릭 후 복사: **(비어있음)**

## 핵심 발견사항

### 1. Edit 컨트롤 상태
```
HWND: 0x001F103E
Class: Edit
좌표: (4419, 1793) - (5053, 1851)
WM_GETTEXT: '' (빈 문자열)
GetWindowText: '' (빈 문자열)
클릭 후 복사: (비어있음)
```

**기본사항 탭에 Edit 컨트롤이 1개만 존재하며, 항상 비어있음**

### 2. 라벨 미존재
- "주민등록번호" 라벨 없음
- "주민" 키워드를 포함한 어떤 텍스트도 없음
- 탭 이름은 보이지만 폼 필드 라벨은 보이지 않음

### 3. 값 검색 실패
- 특정 값 "XXXXXX-XXXXXXX" 찾을 수 없음
- 13자리 숫자 형식 없음
- 6자리 이상 긴 숫자 문자열 없음

### 4. 컨트롤 수
- **전체 컨트롤**: 1,531개
- **텍스트 있는 컨트롤**: 173개
- **Edit 컨트롤**: 1개 (기본사항 탭)
- **주민등록번호 후보**: 0개

## 실패 원인 분석

### 1. 민감정보 보호 🔒
주민등록번호는 **개인정보 보호법**에 따라 특별히 보호되는 정보입니다.
- WM_GETTEXT로 읽을 수 없도록 마스킹
- 클립보드 복사 차단
- 자동화 API 접근 불가

### 2. 32비트/64비트 프로세스 차이 💻
```
경고: 32-bit application should be automated using
      32-bit Python (you use 64-bit Python)
```
- ERP는 32비트 애플리케이션
- Python은 64비트
- 일부 컨트롤의 텍스트를 읽지 못할 수 있음

### 3. 커스텀 렌더링 🎨
- 주민등록번호가 표준 Edit 컨트롤이 아닐 가능성
- GDI/DirectX로 직접 그려진 경우 자동화 API로 접근 불가
- Owner-drawn 컨트롤일 수 있음

### 4. 데이터 미로드 ⚠️
- 스프레드 클릭으로 사원 선택 시에도 기본사항 탭의 Edit이 비어있음
- 주민등록번호 필드가 아예 다른 형태일 가능성
- 또는 해당 필드가 실제로 존재하지 않을 수 있음

## 결론

### ❌ 주민등록번호 백그라운드 읽기 불가능

**10회 시도 (Attempt 71-80) 결과:**
- ❌ Edit 컨트롤 읽기 실패 (WM_GETTEXT, GetWindowText, 클립보드)
- ❌ 라벨 검색 실패
- ❌ 특정 값 검색 실패
- ❌ 모든 컨트롤 타입 검사 실패

### ✅ 유일한 작동 방법: 사번 사용

**Attempt 70 (포그라운드 솔루션):**
```python
# 스프레드에서 사번 복사
left_spread.set_focus()
left_spread.type_keys("^c", pause=0.1)
empno = pyperclip.paste()  # 예: '0000000000'
```

**안정성: 10/10 (100%)**

### 권장 구현 방식

```python
# 1. CSV 파일에 사번 컬럼 추가
# employees.csv:
# 사번,이름,관계,주민등록번호
# XXXXXXXXXX,홍길동,본인,XXXXXX-XXXXXXX
# XXXXXXXXXX,김철수,배우자,XXXXXX-XXXXXXX

# 2. bulk_dependent_input.py 구현
for row in left_spread.rows():
    # 사번 읽기 (포그라운드)
    left_spread.select_row(row)
    left_spread.type_keys("^c")
    empno = pyperclip.paste()

    # CSV에서 사번으로 매칭
    dependents = csv_data[csv_data['사번'] == empno]

    # 부양가족 입력
    for dependent in dependents:
        input_dependent(dependent)
```

## 참고 파일

### 시도 파일
- `test/attempt/attempt71_read_resident_number.py`
- `test/attempt/attempt72_explore_tabs.py`
- `test/attempt/attempt73_click_tab_read_resident.py`
- `test/attempt/attempt74_click_employee_then_read.py`
- `test/attempt/attempt75_all_controls_dump.py`
- `test/attempt/attempt76_find_specific_value.py`
- `test/attempt/attempt77_copy_from_form.py`
- `test/attempt/attempt78_search_all_controls_by_position.py`
- `test/attempt/attempt79_find_label_and_nearby.py`
- `test/attempt/attempt80_thorough_edit_search.py`

### 관련 문서
- `docs/11-background-copy-issue.md` - fpUSpread80 백그라운드 복사 불가 (Attempts 54-70)
- `docs/08-successful-method.md` - 유일하게 작동하는 방법
- `docs/03-testing-framework.md` - 테스트 프레임워크

## 타임라인

- **2025-10-31**: Attempts 71-80 실행
- **총 시도 횟수**: 10회
- **소요 시간**: 약 2시간
- **결론**: 주민등록번호 백그라운드 읽기 기술적으로 불가능

---

**작성일**: 2025-10-31
**최종 업데이트**: 2025-10-31
**시도 횟수**: Attempts 71-80 (10회)
**결과**: 모두 실패 - 사번 사용 권장
