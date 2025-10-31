# 부양가족 대량 입력 구현 계획

## 개요

CSV 파일에서 부양가족 데이터를 읽어 왼쪽 사원 목록과 매칭하여 자동으로 입력하는 스크립트

## 전체 아키텍처

```
┌─────────────────┐
│  CSV 파일       │
│  (테스트데이터) │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  CSVReader      │
│  - 데이터 읽기  │
│  - 사원별 그룹화│
└────────┬────────┘
         │
         ▼
┌──────────────────────────┐
│  BulkDependentInput      │
│                          │
│  1. 왼쪽 목록 스캔       │
│  2. 사번/이름 읽기       │
│  3. CSV 데이터 매칭      │
│  4. 부양가족 입력        │
│  5. 결과 집계 & 로깅     │
└──────────────────────────┘
         │
         ▼
┌─────────────────┐
│  사원등록 프로그램│
│  (부양가족정보) │
└─────────────────┘
```

## 클래스 구조

### BulkDependentInput

```python
class BulkDependentInput:
    """부양가족 대량 입력 자동화"""

    # 초기화
    def __init__(self, csv_path: str)

    # 내부 헬퍼
    def _find_left_spread(self) -> object
    def _find_dependent_dialog(self) -> object
    def log(self, level: str, message: str)

    # 왼쪽 목록 조작
    def count_employees(self) -> int
    def read_current_employee(self) -> tuple[str, str]

    # 부양가족 입력
    def input_dependent(self, dep: DependentData) -> bool
    def process_current_employee(self) -> dict

    # 실행
    def run(self, limit: int, skip: int, dry_run: bool) -> dict
    def _summarize_results(self, results: list) -> dict
```

## 주요 메서드 상세

### 1. __init__(csv_path)

**목적**: 초기화 및 연결

**동작**:
1. CSVReader로 CSV 파일 읽기
2. 사원별로 데이터 그룹화 (dict)
3. pywinauto로 "사원등록" 윈도우 연결
4. 왼쪽 스프레드 찾기 (fpUSpread80)
5. 로그 파일 경로 생성

**반환**: None

---

### 2. _find_left_spread()

**목적**: 왼쪽 사원 목록 스프레드 찾기

**동작**:
1. fpUSpread80 컨트롤 모두 찾기
2. X 좌표로 정렬
3. 가장 왼쪽 것 반환

**반환**: pywinauto wrapper object

---

### 3. count_employees()

**목적**: 왼쪽 목록의 총 사원 수 확인

**동작**:
1. 왼쪽 스프레드 포커스
2. HOME 키로 처음으로
3. 루프:
   - Ctrl+C로 현재 셀 복사
   - 클립보드 읽기
   - 비어있으면 중단
   - 카운트++
   - DOWN 키
4. HOME으로 다시 처음으로

**반환**: int (사원 수)

**에러 처리**:
- 최대 1000회 반복 (무한루프 방지)
- 빈 값을 만나면 중단

---

### 4. read_current_employee()

**목적**: 현재 선택된 사원의 사번과 이름 읽기

**동작**:
1. 왼쪽 스프레드 포커스 확인
2. HOME 키 (행의 처음 = 사번 컬럼)
3. Ctrl+C → 클립보드 읽기 → 사번
4. RIGHT 키 (다음 컬럼 = 이름)
5. Ctrl+C → 클립보드 읽기 → 이름
6. HOME 키 (다시 처음으로)

**반환**: tuple[str, str] (사번, 이름)

**타이밍**:
- 각 키 입력 후 0.2초 대기
- 클립보드 복사 후 0.2초 대기

**클립보드 전략**:
```python
import pyperclip

# 복사
self.dlg.type_keys("^c", pause=0.1)
time.sleep(0.2)

# 읽기
value = pyperclip.paste().strip()
```

---

### 5. input_dependent(dep: DependentData)

**목적**: 부양가족 한 명의 데이터 입력

**입력 순서**:
```
1. 관계코드 (dep.relationship_code)
   → TAB (0.2초 대기)

2. 이름 (dep.name)
   → TAB (0.2초 대기)

3. 내/외국인 (dep.nationality)
   → TAB (0.2초 대기)

4. 주민등록번호 (dep.id_number)
   → ENTER (0.3초 대기)
```

**반환**: bool (성공 여부)

**에러 처리**:
- Exception 발생시 False 반환
- 로그 기록

---

### 6. process_current_employee()

**목적**: 현재 사원의 부양가족 전체 처리

**동작**:
1. 사번/이름 읽기
   - 실패시: 에러 반환

2. CSV 데이터 찾기
   - 없으면: skip 반환

3. 부양가족 필터링
   - 관계코드 != '0' (본인 제외)
   - 없으면: skip 반환

4. 각 부양가족 입력
   - 성공/실패 카운트
   - 개별 로그 기록

5. 결과 반환

**반환**:
```python
{
    'status': 'success' | 'skip' | 'error',
    'employee_no': str,
    'employee_name': str,
    'total': int,
    'success': int,
    'reason': str  # skip/error인 경우
}
```

---

### 7. run(limit, skip, dry_run)

**목적**: 전체 자동화 실행

**파라미터**:
- `limit`: 처리할 최대 사원 수 (None = 전체)
- `skip`: 처음 건너뛸 사원 수 (0 = 처음부터)
- `dry_run`: True면 실제 입력 안함 (테스트용)

**동작**:
```
1. 초기화
   - 로그 헤더 출력
   - CSV 정보 출력

2. 왼쪽 목록 스캔
   - 포커스
   - 사원 수 확인

3. 처리 범위 계산
   - start_idx = skip
   - end_idx = skip + limit (또는 total)

4. 시작 위치로 이동
   - HOME
   - DOWN × skip

5. 각 사원 반복 처리
   - dry_run: 읽기만
   - 아니면: process_current_employee()
   - DOWN으로 다음 사원

6. 결과 집계 및 출력
```

**반환**: dict (요약 결과)

---

### 8. _summarize_results(results)

**목적**: 결과 리스트 집계

**반환**:
```python
{
    'processed': int,       # 처리한 사원 수
    'success': int,         # 성공한 사원 수
    'skipped': int,         # 건너뛴 사원 수
    'failed': int,          # 실패한 사원 수
    'total_dependents': int # 입력한 부양가족 수
}
```

## 로깅 전략

### 로그 레벨

- **INFO**: 진행 상황
- **SUCCESS**: 성공한 작업
- **WARNING**: 주의 사항
- **ERROR**: 실패한 작업

### 로그 포맷

```
[2025-10-30 19:05:30] INFO: 메시지
```

### 로그 파일

```
logs/bulk_input_20251030_190530.log
```

### 주요 로그 포인트

```python
# 시작
"=== 부양가족 대량 입력 시작 ==="
"CSV 파일: 테스트 데이터.csv"
"CSV 사원 수: 75"
"왼쪽 목록 사원 수: 75"
"처리 범위: 1번째 ~ 10번째 사원"

# 사원 처리
"[1/75]"
"처리 시작: 2021122920 (강경민)"
"  → 부양가족 2명 입력 시작"
"    ✓ 유귀열 (관계: 1)"
"    ✓ 강창구 (관계: 1)"
"  → 완료: 2/2"

# 건너뛰기
"처리 시작: 2022021605 (박세연)"
"  → 부양가족 없음, 건너뜀"

# 에러
"사원 정보 읽기 실패: ..."

# 결과
"=== 결과 요약 ==="
"처리 사원: 10"
"성공: 8"
"건너뜀: 2"
"실패: 0"
"입력 부양가족: 26"
```

## 클립보드 읽기 상세

### 라이브러리

```python
import pyperclip
```

### 설치

```bash
uv add pyperclip
```

### 사용 패턴

```python
# 1. 포커스 확인
self.left_spread.set_focus()
time.sleep(0.2)

# 2. 셀 이동 (필요시)
self.dlg.type_keys("{HOME}", pause=0.1)
time.sleep(0.2)

# 3. 복사
self.dlg.type_keys("^c", pause=0.1)
time.sleep(0.2)

# 4. 읽기
value = pyperclip.paste().strip()

# 5. 검증
if not value:
    # 빈 값 처리
```

### 주의사항

1. **타이밍**: 각 단계마다 충분한 대기 시간
2. **포커스**: 복사 전 반드시 포커스 확인
3. **trim**: `strip()`으로 공백 제거
4. **멀티스레드**: 클립보드는 단일 스레드에서만 사용

## 실행 옵션

### 명령줄 인터페이스

```bash
# 기본 실행
python bulk_dependent_input.py --csv "테스트 데이터.csv"

# 처음 5명만
python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 5

# 10번째부터 5명
python bulk_dependent_input.py --csv "테스트 데이터.csv" --skip 10 --limit 5

# 드라이런 (테스트)
python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 3 --dry-run
```

### argparse 설정

```python
parser = argparse.ArgumentParser(description='부양가족 대량 입력')
parser.add_argument('--csv', required=True, help='CSV 파일 경로')
parser.add_argument('--limit', type=int, help='처리할 사원 수 제한')
parser.add_argument('--skip', type=int, default=0, help='건너뛸 사원 수')
parser.add_argument('--dry-run', action='store_true', help='실제 입력 없이 테스트')
```

## 에러 처리 전략

### 1. 연결 실패

```python
try:
    self.dlg = self.app.connect(title_re=".*사원등록.*")
except:
    print("ERROR: 사원등록 프로그램을 찾을 수 없습니다")
    sys.exit(1)
```

### 2. 사원 읽기 실패

```python
try:
    emp_no, emp_name = self.read_current_employee()
except Exception as e:
    self.log("ERROR", f"사원 정보 읽기 실패: {e}")
    return {'status': 'error', 'reason': 'read_failed'}
```

### 3. 입력 실패

```python
try:
    self.input_dependent(dep)
except Exception as e:
    self.log("ERROR", f"{dep.name} 입력 실패: {e}")
    # 다음 부양가족 계속 진행
```

### 4. 전체 실패

```python
try:
    bulk.run(...)
except KeyboardInterrupt:
    print("\n\n중단됨 (Ctrl+C)")
except Exception as e:
    print(f"\n\n치명적 오류: {e}")
    import traceback
    traceback.print_exc()
```

## 테스트 계획

### 1단계: 드라이런

```bash
# 처음 3명만 읽기 테스트
python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 3 --dry-run
```

**확인사항**:
- 사원 수 카운트 정확한가?
- 사번/이름 읽기 정확한가?
- CSV 매칭 정확한가?

### 2단계: 소량 입력

```bash
# 처음 1명만 실제 입력
python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 1
```

**확인사항**:
- 부양가족 데이터 정확히 입력되는가?
- 다음 행으로 제대로 이동하는가?
- 로그가 정확한가?

### 3단계: 중간 규모

```bash
# 처음 10명
python bulk_dependent_input.py --csv "테스트 데이터.csv" --limit 10
```

### 4단계: 전체

```bash
# 전체 실행
python bulk_dependent_input.py --csv "테스트 데이터.csv"
```

## 디렉토리 구조

```
newgen-erp-macro/
├── src/
│   ├── csv_reader.py           # 기존
│   └── bulk_dependent_input.py # 새로 작성
├── logs/                        # 새로 생성
│   └── bulk_input_*.log
├── 테스트 데이터.csv
└── docs/
    └── bulk-input-implementation.md  # 이 문서
```

## 의존성

```toml
# pyproject.toml
[project]
dependencies = [
    "pywinauto>=0.6.8",
    "pywin32>=306",
    "pillow>=10.0.0",
    "mss>=9.0.0",
    "pyperclip>=1.8.2",  # 추가
]
```

## 다음 단계

1. ✅ 구현 계획 문서화 (이 문서)
2. 🔄 뼈대 코드 작성
3. ⏳ 클립보드 읽기 테스트
4. ⏳ 드라이런 모드 테스트
5. ⏳ 소량 입력 테스트
6. ⏳ 전체 입력 테스트
7. ⏳ 에러 처리 강화
8. ⏳ 문서 업데이트
