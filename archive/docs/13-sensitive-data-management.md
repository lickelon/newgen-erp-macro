# 민감정보 관리 방법

## 개요

프로젝트에서 테스트 시 필요한 민감정보(사번, 주민등록번호 등)를 안전하게 관리하는 방법을 설명합니다.

## ❌ 잘못된 방법: 하드코딩

```python
# ❌ 이렇게 하지 마세요!
def run(dlg, capture_func):
    target_empno = "0000000000"  # 실제 사번 하드코딩
    target_resident = "123456-1234567"  # 실제 주민등록번호 하드코딩

    # ... 테스트 코드 ...
```

**문제점:**
- 퍼블릭 레포지토리에 민감정보가 노출됨
- 개인정보 보호법 위반 가능성
- 보안 취약점

## ✅ 올바른 방법: 설정 파일 사용

### 1. 설정 파일 구조

```
newgen-erp-macro/
├── test_config.example.json  # 템플릿 (git 추적)
├── test_config.json           # 실제 데이터 (git 무시)
├── .gitignore                 # test_config.json 포함
└── test/
    └── config.py              # 설정 로더
```

### 2. 설정 파일 생성

**Step 1**: 템플릿 복사

```bash
copy test_config.example.json test_config.json
```

**Step 2**: `test_config.json`에 실제 값 입력

```json
{
  "test_data": {
    "employee": {
      "empno": "0000000000",
      "name": "홍길동"
    },
    "resident_numbers": {
      "with_hyphen": "123456-1234567",
      "without_hyphen": "12345612345567"
    }
  }
}
```

**Step 3**: `.gitignore` 확인 (이미 설정됨)

```gitignore
# Test configuration (sensitive data)
test_config.json
test_config.*.json
.env
.env.local
```

### 3. 코드에서 사용하기

#### 방법 1: TestConfig 클래스 직접 사용

```python
from test.config import get_test_config

def run(dlg, capture_func):
    try:
        config = get_test_config()

        # 설정에서 데이터 로드
        empno = config.empno
        resident = config.resident_number_with_hyphen

        print(f"사번: {empno}")
        print(f"주민등록번호: {resident}")

    except FileNotFoundError as e:
        return {"success": False, "message": str(e)}
```

#### 방법 2: 편의 함수 사용

```python
from test.config import get_empno, get_resident_number

def run(dlg, capture_func):
    empno = get_empno()
    resident = get_resident_number(with_hyphen=True)

    # ... 테스트 코드 ...
```

#### 방법 3: 점 표기법으로 접근

```python
from test.config import get_test_config

config = get_test_config()

# 중첩된 값 접근
empno = config.get("test_data.employee.empno")
name = config.get("test_data.employee.name")
```

## API 문서

### TestConfig 클래스

#### 메서드

##### `__init__(config_path=None)`
설정 파일 로드

- **config_path**: 설정 파일 경로 (None이면 자동으로 `test_config.json` 찾기)

##### `get(key_path, default=None)`
점 표기법으로 설정값 접근

- **key_path**: `"test_data.employee.empno"` 형식
- **default**: 키가 없을 때 반환할 기본값
- **Returns**: 설정값

#### 프로퍼티

##### `empno`
테스트용 사번 반환

```python
config = get_test_config()
empno = config.empno  # "0000000000"
```

##### `employee_name`
테스트용 사원 이름 반환

```python
name = config.employee_name  # "홍길동"
```

##### `resident_number_with_hyphen`
하이픈 포함 주민등록번호 반환

```python
resident = config.resident_number_with_hyphen  # "123456-1234567"
```

##### `resident_number_without_hyphen`
하이픈 없는 주민등록번호 반환

```python
resident = config.resident_number_without_hyphen  # "12345612345567"
```

### 편의 함수

#### `get_test_config()`
싱글톤 설정 인스턴스 반환

```python
from test.config import get_test_config

config = get_test_config()
```

#### `get_empno()`
테스트용 사번 반환

```python
from test.config import get_empno

empno = get_empno()
```

#### `get_resident_number(with_hyphen=True)`
테스트용 주민등록번호 반환

- **with_hyphen**: True이면 하이픈 포함, False면 하이픈 없음

```python
from test.config import get_resident_number

resident1 = get_resident_number(with_hyphen=True)   # "123456-1234567"
resident2 = get_resident_number(with_hyphen=False)  # "12345612345567"
```

## 예제: attempt82_config_example.py

전체 예제는 `test/attempt/attempt82_config_example.py`를 참고하세요.

```python
from test.config import get_test_config

def run(dlg, capture_func):
    try:
        config = get_test_config()

        target_with_hyphen = config.resident_number_with_hyphen
        target_without_hyphen = config.resident_number_without_hyphen
        empno = config.empno

        print(f"✓ 설정 파일 로드 성공!")
        print(f"  사번: '{empno}'")
        print(f"  주민등록번호 (하이픈 포함): '{target_with_hyphen}'")

    except FileNotFoundError as e:
        return {"success": False, "message": str(e)}

    # ... 테스트 코드 ...
```

## 에러 처리

### FileNotFoundError

설정 파일이 없을 때 발생합니다.

```python
try:
    config = get_test_config()
except FileNotFoundError as e:
    print(f"❌ {e}")
    print("test_config.example.json을 test_config.json으로 복사하세요.")
```

**에러 메시지 예시:**
```
테스트 설정 파일이 없습니다: C:\...\test_config.json
test_config.example.json을 test_config.json으로 복사하고 실제 값을 입력하세요.
```

## 테스트

설정 로더 테스트:

```bash
python -m test.config
```

**출력 예시:**
```
설정 로드 성공!
사번: 0000000000
이름: 홍길동
주민등록번호 (하이픈 포함): 123456-1234567
주민등록번호 (하이픈 없음): 12345612345567
```

## 보안 체크리스트

- [x] `test_config.json`이 `.gitignore`에 포함되어 있는가?
- [x] `test_config.example.json`은 마스킹된 예시값만 포함하는가?
- [x] 코드에 하드코딩된 민감정보가 없는가?
- [x] CSV/Excel 테스트 데이터도 `.gitignore`에 포함되어 있는가?
- [x] 커밋 전 `git status`로 민감정보 파일이 포함되지 않았는지 확인했는가?

## 추가 설정 확장

새로운 테스트 데이터 추가:

```json
{
  "test_data": {
    "employee": {
      "empno": "0000000000",
      "name": "홍길동"
    },
    "resident_numbers": {
      "with_hyphen": "123456-1234567",
      "without_hyphen": "12345612345567"
    },
    "department": {
      "code": "DEV001",
      "name": "개발팀"
    }
  }
}
```

코드에서 접근:

```python
config = get_test_config()
dept_code = config.get("test_data.department.code")
dept_name = config.get("test_data.department.name")
```

## 관련 파일

- `test_config.example.json` - 설정 파일 템플릿
- `test/config.py` - 설정 로더 유틸리티
- `test/attempt/attempt82_config_example.py` - 사용 예제
- `.gitignore` - 민감정보 파일 제외 설정

---

**작성일**: 2025-10-31
**최종 업데이트**: 2025-10-31
**관련 이슈**: 민감정보 보호, 개인정보 보호법 준수
