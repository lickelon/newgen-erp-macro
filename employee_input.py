"""
사원등록 프로그램 - 직원 정보 입력 자동화

SPR32DU80EditHScroll 컨트롤에 직원 정보를 자동으로 입력합니다.
마우스 커서를 움직이지 않고 윈도우 메시지만 사용합니다.

성공 방법 (Attempt 09):
1. SPR32DU80EditHScroll 컨트롤 찾기
2. WM_SETTEXT로 텍스트 설정
3. EN_CHANGE 알림을 부모에게 전송
4. Enter 키 전송
"""
import time
import win32api
import win32con
import win32gui
from pywinauto import application


class EmployeeInput:
    """직원 정보 입력 자동화 클래스"""

    def __init__(self, window_title="사원등록"):
        """
        Args:
            window_title: 타겟 윈도우 제목 (기본: "사원등록")
        """
        self.window_title = window_title
        self.app = None
        self.dlg = None
        self.edit_controls = []

        # 메시지 상수
        self.EN_CHANGE = 0x0300
        self.WM_COMMAND = 0x0111

    def connect(self):
        """사원등록 프로그램에 연결"""
        try:
            self.app = application.Application(backend="win32")
            self.app.connect(title=self.window_title)
            self.dlg = self.app.window(title=self.window_title)
            return True
        except Exception as e:
            raise Exception(f"사원등록 윈도우를 찾을 수 없습니다: {e}")

    def find_edit_controls(self):
        """SPR32DU80EditHScroll 컨트롤 찾기"""
        self.edit_controls = []

        for ctrl in self.dlg.descendants():
            try:
                if "SPR32DU80EditHScroll" in ctrl.class_name():
                    self.edit_controls.append(ctrl)
            except:
                pass

        if not self.edit_controls:
            raise Exception("SPR32DU80EditHScroll 컨트롤을 찾을 수 없습니다")

        return len(self.edit_controls)

    def input_field(self, ctrl, text, send_enter=True):
        """
        단일 필드에 텍스트 입력

        Args:
            ctrl: pywinauto 컨트롤 객체
            text: 입력할 텍스트
            send_enter: Enter 키 전송 여부 (기본: True)

        Returns:
            bool: 성공 여부
        """
        try:
            hwnd = ctrl.handle

            # 1. WM_SETTEXT로 텍스트 설정
            win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
            time.sleep(0.05)

            # 2. EN_CHANGE 알림을 부모에게 전송
            try:
                parent_hwnd = win32gui.GetParent(hwnd)
                if parent_hwnd:
                    ctrl_id = win32api.GetWindowLong(hwnd, win32con.GWL_ID)
                    wparam = (self.EN_CHANGE << 16) | ctrl_id
                    win32api.SendMessage(parent_hwnd, self.WM_COMMAND, wparam, hwnd)
            except:
                pass

            time.sleep(0.05)

            # 3. Enter 키 전송 (선택적)
            if send_enter:
                win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                time.sleep(0.02)
                win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                time.sleep(0.05)

            # 4. 결과 확인
            result_text = ctrl.window_text()
            return result_text == text

        except Exception as e:
            print(f"입력 오류: {e}")
            return False

    def input_employee(self, employee_no=None, id_number=None, name=None):
        """
        직원 정보 입력

        Args:
            employee_no: 사번 (첫 번째 Edit 컨트롤)
            id_number: 주민번호 (두 번째 Edit 컨트롤)
            name: 성명 (세 번째 Edit 컨트롤)

        Returns:
            dict: 입력 결과
                {
                    "success": bool,
                    "results": [
                        {"field": "사번", "success": bool, "value": str},
                        ...
                    ]
                }
        """
        # 컨트롤 찾기
        if not self.edit_controls:
            self.find_edit_controls()

        inputs = [
            ("사번", employee_no),
            ("주민번호", id_number),
            ("성명", name),
        ]

        results = []
        success_count = 0

        for idx, (field_name, value) in enumerate(inputs):
            if idx >= len(self.edit_controls):
                results.append({
                    "field": field_name,
                    "success": False,
                    "message": "컨트롤을 찾을 수 없음"
                })
                continue

            if value is None:
                results.append({
                    "field": field_name,
                    "success": False,
                    "message": "값이 제공되지 않음"
                })
                continue

            ctrl = self.edit_controls[idx]
            success = self.input_field(ctrl, value)

            results.append({
                "field": field_name,
                "success": success,
                "value": value
            })

            if success:
                success_count += 1

        return {
            "success": success_count > 0,
            "total": len([v for _, v in inputs if v is not None]),
            "success_count": success_count,
            "results": results
        }

    def get_current_values(self):
        """현재 Edit 컨트롤의 값들 조회"""
        if not self.edit_controls:
            self.find_edit_controls()

        values = []
        field_names = ["사번", "주민번호", "성명"]

        for idx, ctrl in enumerate(self.edit_controls):
            try:
                field_name = field_names[idx] if idx < len(field_names) else f"필드{idx+1}"
                value = ctrl.window_text()
                values.append({
                    "field": field_name,
                    "value": value,
                    "hwnd": f"0x{ctrl.handle:08X}"
                })
            except Exception as e:
                values.append({
                    "field": field_name,
                    "value": None,
                    "error": str(e)
                })

        return values


def main():
    """사용 예제"""
    import sys

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    print("=" * 70)
    print("직원 정보 입력 자동화")
    print("=" * 70)

    # 1. 연결
    print("\n[1/4] 사원등록 프로그램 연결...")
    try:
        emp_input = EmployeeInput()
        emp_input.connect()
        print("✓ 연결 성공")
    except Exception as e:
        print(f"✗ 연결 실패: {e}")
        return

    # 2. 기본사항 탭 선택
    print("\n[2/4] 기본사항 탭 선택...")
    from tab_automation import TabAutomation
    tab_auto = TabAutomation()
    tab_auto.connect()
    tab_auto.select_tab("기본사항")
    time.sleep(0.5)
    print("✓ 기본사항 탭 선택됨")

    # 3. Edit 컨트롤 찾기
    print("\n[3/4] Edit 컨트롤 찾기...")
    try:
        count = emp_input.find_edit_controls()
        print(f"✓ {count}개 컨트롤 발견")
    except Exception as e:
        print(f"✗ 컨트롤 찾기 실패: {e}")
        return

    # 현재 값 확인
    print("\n현재 값:")
    for item in emp_input.get_current_values():
        print(f"  • {item['field']}: \"{item['value']}\"")

    # 4. 데이터 입력
    print("\n[4/4] 직원 정보 입력...")
    test_data = {
        "employee_no": "2025001",
        "id_number": "900101-1234567",
        "name": "홍길동"
    }
    print(f"입력 데이터: {test_data}")

    result = emp_input.input_employee(**test_data)

    print(f"\n결과: {result['success_count']}/{result['total']}개 성공")
    for item in result['results']:
        status = "✅" if item['success'] else "❌"
        print(f"  {status} {item['field']}: {item.get('value', 'N/A')}")

    # 입력 후 값 확인
    print("\n입력 후 값:")
    for item in emp_input.get_current_values():
        print(f"  • {item['field']}: \"{item['value']}\"")

    print("\n" + "=" * 70)
    if result['success']:
        print("✅ 직원 정보 입력 완료!")
    else:
        print("⚠️  일부 입력 실패")
    print("=" * 70)


if __name__ == "__main__":
    main()
