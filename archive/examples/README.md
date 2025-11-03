# 자동화 예제

이 폴더는 성공한 자동화 방법들의 실제 작동 예제를 포함합니다.

## 포함된 예제

### 1. example_employee_selection.py
**원본**: attempt53_select_employee.py
**성공률**: ✅ 100%
**기능**: 좌표 없이 왼쪽 스프레드시트에서 사원 선택

**핵심 방법**:
```python
# 왼쪽 스프레드 찾기 (X 좌표로 정렬)
spreads.sort(key=lambda s: s.rectangle().left)
left_spread = spreads[0]

# 포커스 + 키보드
left_spread.set_focus()
dlg.type_keys("{HOME}")  # 첫 번째로
dlg.type_keys("{DOWN}")  # 다음 사원
```

### 2. example_tab_switching.py
**원본**: attempt52_dialog_combinations.py
**성공률**: ✅ 100%
**기능**: 좌표 없이 탭 전환 (기본사항 ↔ 부양가족정보)

**핵심 방법**:
```python
# ShowWindow로 Dialog 숨기고 보이기
win32gui.ShowWindow(other_dialog.handle, win32con.SW_HIDE)
win32gui.ShowWindow(target_dialog.handle, win32con.SW_SHOW)
```

**중요**: 관리자 권한 필요 (UIPI 보안)

### 3. example_data_input.py
**원본**: attempt43_dlg_type_keys.py
**성공률**: ✅ 100%
**기능**: 좌표 없이 부양가족 데이터 입력

**핵심 방법**:
```python
# Dialog 레벨에서 type_keys 사용
dlg.type_keys("4", with_spaces=False, pause=0.05)
dlg.type_keys("{TAB}", pause=0.05)
dlg.type_keys("강부양", with_spaces=False, pause=0.05)
dlg.type_keys("{ENTER}", pause=0.05)
```

## 사용법

각 예제를 단독 실행:
```bash
# 관리자 권한 터미널에서
uv run python examples/example_employee_selection.py
uv run python examples/example_tab_switching.py
uv run python examples/example_data_input.py
```

## 완전한 자동화

전체 워크플로우는 [docs/automation-guide.md](../docs/automation-guide.md)를 참조하세요.

## 테스트 이력

이 예제들은 53회의 시도 끝에 발견된 성공적인 방법들입니다:
- Attempt 01-42: 다양한 방법 시도 (대부분 실패)
- Attempt 43: 데이터 입력 성공
- Attempt 47-51: 탭 전환 실패
- Attempt 52: 탭 전환 성공
- Attempt 53: 사원 선택 성공
