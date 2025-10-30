"""
사원등록 프로그램 - 직원 정보 입력 자동화

fpUSpread80 스프레드시트 컨트롤에 직원 정보를 자동으로 입력합니다.
마우스 커서를 움직이지 않고 윈도우 메시지만 사용합니다.

확립된 방법 (Attempt 18):
1. fpUSpread80 Spread #2 (왼쪽 직원 목록) 찾기
2. WM_LBUTTONDOWN/UP로 셀 클릭하여 선택
3. WM_CHAR로 각 문자 입력
4. VK_RETURN으로 Enter 키 전송

좌표 매핑:
- x=50:  사번 (Employee Number)
- x=100: 성명 (Name)
- x=200: 주민번호 (ID Number)
- x=320: 나이 (Age)
- y=30:  새 행 삽입 위치
"""
import time
import win32api
import win32con
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
        self.spread_hwnd = None

        # 확립된 좌표 매핑 (Spread #2 - 왼쪽 직원 목록)
        self.field_coords = {
            "사번": {"x": 50, "y": 30},
            "성명": {"x": 100, "y": 30},
            "주민번호": {"x": 200, "y": 30},
            "나이": {"x": 320, "y": 30},
        }

    def connect(self):
        """사원등록 프로그램에 연결"""
        try:
            self.app = application.Application(backend="win32")
            self.app.connect(title=self.window_title)
            self.dlg = self.app.window(title=self.window_title)
            return True
        except Exception as e:
            raise Exception(f"사원등록 윈도우를 찾을 수 없습니다: {e}")

    def find_spread_control(self):
        """
        fpUSpread80 Spread #2 (왼쪽 직원 목록) 찾기

        Returns:
            int: Spread HWND, 또는 None
        """
        spread_controls = []
        for ctrl in self.dlg.descendants():
            try:
                if ctrl.class_name() == "fpUSpread80":
                    spread_controls.append(ctrl)
            except:
                pass

        if len(spread_controls) < 3:
            raise Exception(f"fpUSpread80 컨트롤 부족 (발견: {len(spread_controls)}, 필요: 3)")

        # Spread #2 = 왼쪽 직원 목록
        self.spread_hwnd = spread_controls[2].handle
        return self.spread_hwnd

    def input_to_cell(self, x, y, text, label=None):
        """
        fpUSpread80의 특정 셀에 텍스트 입력

        Args:
            x: 셀 X 좌표
            y: 셀 Y 좌표
            text: 입력할 텍스트
            label: 필드 이름 (로깅용, 선택적)

        Returns:
            bool: 성공 여부
        """
        if not self.spread_hwnd:
            raise Exception("Spread 컨트롤이 초기화되지 않았습니다. find_spread_control()을 먼저 호출하세요.")

        try:
            if label:
                print(f"  {label}: \"{text}\" at ({x},{y})")

            # 1. 셀 클릭하여 선택
            lparam = win32api.MAKELONG(x, y)
            win32api.SendMessage(self.spread_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            time.sleep(0.03)
            win32api.SendMessage(self.spread_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            time.sleep(0.2)

            # 2. 각 문자 입력
            for char in text:
                win32api.SendMessage(self.spread_hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.015)

            # 3. Enter 키 전송
            win32api.SendMessage(self.spread_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            time.sleep(0.02)
            win32api.SendMessage(self.spread_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            time.sleep(0.5)

            return True

        except Exception as e:
            print(f"    ✗ 입력 오류: {e}")
            return False

    def input_employee(self, employee_no=None, name=None, id_number=None, age=None):
        """
        직원 정보 입력

        Args:
            employee_no: 사번
            name: 성명
            id_number: 주민번호
            age: 나이

        Returns:
            dict: 입력 결과
                {
                    "success": bool,
                    "total": int,
                    "success_count": int,
                    "results": [
                        {"field": "사번", "success": bool, "value": str},
                        ...
                    ]
                }
        """
        # Spread 컨트롤 찾기
        if not self.spread_hwnd:
            self.find_spread_control()

        # 입력할 데이터 준비
        inputs = [
            ("사번", employee_no, self.field_coords["사번"]),
            ("성명", name, self.field_coords["성명"]),
            ("주민번호", id_number, self.field_coords["주민번호"]),
            ("나이", age, self.field_coords["나이"]),
        ]

        results = []
        success_count = 0

        print("\n직원 정보 입력 중...")
        for field_name, value, coords in inputs:
            if value is None:
                results.append({
                    "field": field_name,
                    "success": False,
                    "message": "값이 제공되지 않음"
                })
                continue

            success = self.input_to_cell(
                x=coords["x"],
                y=coords["y"],
                text=str(value),
                label=field_name
            )

            results.append({
                "field": field_name,
                "success": success,
                "value": value
            })

            if success:
                success_count += 1
                print(f"    ✓ {field_name} 입력 완료")

            time.sleep(0.3)  # 각 필드 입력 후 대기

        return {
            "success": success_count > 0,
            "total": len([v for _, v, _ in inputs if v is not None]),
            "success_count": success_count,
            "results": results
        }


def main():
    """사용 예제"""
    import sys

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    print("=" * 70)
    print("직원 정보 입력 자동화 (fpUSpread80 방식)")
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

    # 3. Spread 컨트롤 찾기
    print("\n[3/4] fpUSpread80 컨트롤 찾기...")
    try:
        hwnd = emp_input.find_spread_control()
        print(f"✓ Spread #2 발견 (HWND=0x{hwnd:08X})")
    except Exception as e:
        print(f"✗ 컨트롤 찾기 실패: {e}")
        return

    # 4. 데이터 입력
    print("\n[4/4] 직원 정보 입력...")
    test_data = {
        "employee_no": "2025100",
        "name": "테스트사원",
        "id_number": "920315-1234567",
        "age": "33"
    }
    print(f"입력 데이터: {test_data}")

    result = emp_input.input_employee(**test_data)

    print(f"\n결과: {result['success_count']}/{result['total']}개 성공")
    for item in result['results']:
        status = "✅" if item['success'] else "❌"
        print(f"  {status} {item['field']}: {item.get('value', 'N/A')}")

    print("\n" + "=" * 70)
    if result['success']:
        print("✅ 직원 정보 입력 완료!")
        print("⚠️  화면에서 입력 결과를 확인하세요.")
    else:
        print("⚠️  일부 입력 실패")
    print("=" * 70)


if __name__ == "__main__":
    main()
