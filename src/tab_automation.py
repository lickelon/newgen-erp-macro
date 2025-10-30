"""
사원등록 프로그램 탭 자동화 모듈
Win32 SendMessage를 사용하여 마우스 커서를 움직이지 않고 탭 선택

성공 방법 (2025-10-30):
- 탭 컨트롤 HWND에 직접 WM_LBUTTONDOWN/UP 전송
- 클라이언트 좌표 사용 (탭 컨트롤 내부 기준)
- 마우스 커서 움직임 없음 확인됨
"""
import time
import win32api
import win32con
from pywinauto import application


class TabAutomation:
    """탭 자동화 클래스"""

    # 탭 이름과 X 좌표 매핑 (Y는 항상 15)
    TAB_POSITIONS = {
        "기본사항": 50,
        "부양가족정보": 150,
        "소득자료": 250,
        # 추가 탭이 있다면 여기에 추가
    }

    def __init__(self, window_title="사원등록"):
        """
        Args:
            window_title: 연결할 윈도우 제목
        """
        self.window_title = window_title
        self.app = None
        self.dlg = None
        self.tab_control = None

    def connect(self):
        """사원등록 윈도우에 연결"""
        try:
            self.app = application.Application(backend="win32")
            self.app.connect(title=self.window_title)
            self.dlg = self.app.window(title=self.window_title)

            # 탭 컨트롤 찾기 (부분 매칭 사용 - 프로그램 재시작에도 안정적)
            self.tab_control = None
            descendants = self.dlg.descendants()

            for ctrl in descendants:
                try:
                    # Afx:TabWnd:로 시작하는 클래스명 찾기
                    if ctrl.class_name().startswith("Afx:TabWnd:"):
                        self.tab_control = ctrl
                        break
                except:
                    pass

            if self.tab_control is None:
                raise Exception("탭 컨트롤을 찾을 수 없습니다")

            return True

        except Exception as e:
            raise Exception(f"윈도우 연결 실패: {e}")

    def select_tab(self, tab_name, wait_time=0.5):
        """
        특정 탭을 선택 (마우스 커서 움직이지 않음)

        Args:
            tab_name: 선택할 탭 이름 (예: "부양가족정보")
            wait_time: 탭 선택 후 대기 시간 (초)

        Returns:
            bool: 성공 여부
        """
        if tab_name not in self.TAB_POSITIONS:
            available = ", ".join(self.TAB_POSITIONS.keys())
            raise ValueError(f"알 수 없는 탭: {tab_name}. 사용 가능한 탭: {available}")

        if not self.tab_control:
            raise Exception("먼저 connect()를 호출하세요")

        # 탭 컨트롤 핸들 가져오기
        hwnd = self.tab_control.handle

        # 클릭 좌표
        x = self.TAB_POSITIONS[tab_name]
        y = 15  # 탭 영역의 중앙 Y 좌표

        # LPARAM = MAKELONG(x, y)
        lparam = win32api.MAKELONG(x, y)

        # WM_LBUTTONDOWN
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.1)

        # WM_LBUTTONUP
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(wait_time)

        return True

    def select_tab_by_index(self, index, wait_time=0.5):
        """
        인덱스로 탭 선택 (0부터 시작)

        Args:
            index: 탭 인덱스 (0=기본사항, 1=부양가족정보, 2=소득자료, ...)
            wait_time: 탭 선택 후 대기 시간 (초)

        Returns:
            bool: 성공 여부
        """
        if not self.tab_control:
            raise Exception("먼저 connect()를 호출하세요")

        # 탭 컨트롤 핸들 가져오기
        hwnd = self.tab_control.handle

        # 각 탭의 대략적인 X 좌표 계산
        tab_width = 100  # 탭 하나당 대략적인 너비
        x = 50 + (index * tab_width)
        y = 15

        # LPARAM = MAKELONG(x, y)
        lparam = win32api.MAKELONG(x, y)

        # WM_LBUTTONDOWN
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.1)

        # WM_LBUTTONUP
        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        time.sleep(wait_time)

        return True


# 사용 예제
if __name__ == "__main__":
    import sys

    # UTF-8 출력
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    print("="*70)
    print("탭 자동화 테스트")
    print("="*70)

    # 탭 자동화 객체 생성
    tab_auto = TabAutomation()

    try:
        # 연결
        print("\n✓ 사원등록 윈도우 연결 중...")
        tab_auto.connect()
        print("✓ 연결 성공!\n")

        # 탭 순서대로 선택
        tabs = ["기본사항", "부양가족정보", "소득자료"]

        for tab_name in tabs:
            print(f"'{tab_name}' 탭 선택 중...")
            tab_auto.select_tab(tab_name)
            print(f"✓ '{tab_name}' 탭 선택 완료\n")
            time.sleep(1)

        print("="*70)
        print("✅ 모든 탭 자동 선택 완료!")
        print("="*70)

    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        sys.exit(1)
