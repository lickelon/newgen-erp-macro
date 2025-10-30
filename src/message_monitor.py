"""
Python으로 윈도우 메시지 직접 모니터링

Spy++ 없이 Python만으로 윈도우 메시지를 실시간 캡처합니다.
"""
import sys
import time
import threading
import ctypes
from ctypes import wintypes
import win32con
import win32api
from pywinauto import application

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

# Windows API 상수
WH_CALLWNDPROC = 4
WH_GETMESSAGE = 3

# 메시지 저장용 리스트
captured_messages = []
monitoring = False


class CWPSTRUCT(ctypes.Structure):
    """CallWndProc 구조체"""
    _fields_ = [
        ("lParam", wintypes.LPARAM),
        ("wParam", wintypes.WPARAM),
        ("message", wintypes.UINT),
        ("hwnd", wintypes.HWND),
    ]


def decode_lparam_coords(lparam):
    """LPARAM에서 x, y 좌표 추출"""
    x = lparam & 0xFFFF
    y = (lparam >> 16) & 0xFFFF
    return x, y


def message_to_string(msg_code):
    """메시지 코드를 문자열로 변환"""
    message_names = {
        win32con.WM_LBUTTONDOWN: "WM_LBUTTONDOWN",
        win32con.WM_LBUTTONUP: "WM_LBUTTONUP",
        win32con.WM_NOTIFY: "WM_NOTIFY",
        win32con.WM_PAINT: "WM_PAINT",
        win32con.WM_MOUSEMOVE: "WM_MOUSEMOVE",
        0x130C: "TCM_SETCURSEL",
        win32con.WM_SETFOCUS: "WM_SETFOCUS",
        win32con.WM_KILLFOCUS: "WM_KILLFOCUS",
    }
    return message_names.get(msg_code, f"0x{msg_code:04X}")


def log_message(hwnd, msg, wparam, lparam, prefix=""):
    """메시지 로그 출력"""
    msg_name = message_to_string(msg)

    # 필터링: 관심있는 메시지만 출력
    interesting_messages = [
        win32con.WM_LBUTTONDOWN,
        win32con.WM_LBUTTONUP,
        win32con.WM_NOTIFY,
        0x130C,  # TCM_SETCURSEL
    ]

    if msg not in interesting_messages:
        return

    timestamp = time.strftime("%H:%M:%S")

    # 좌표 메시지인 경우 x, y 추출
    if msg in [win32con.WM_LBUTTONDOWN, win32con.WM_LBUTTONUP]:
        x, y = decode_lparam_coords(lparam)
        print(f"{prefix}[{timestamp}] HWND=0x{hwnd:08X} {msg_name} "
              f"wParam=0x{wparam:08X} lParam=0x{lparam:08X} (x={x}, y={y})")
    else:
        print(f"{prefix}[{timestamp}] HWND=0x{hwnd:08X} {msg_name} "
              f"wParam=0x{wparam:08X} lParam=0x{lparam:08X}")

    # 메시지 저장
    captured_messages.append({
        "timestamp": timestamp,
        "hwnd": hwnd,
        "msg": msg,
        "msg_name": msg_name,
        "wparam": wparam,
        "lparam": lparam,
    })


class MessageMonitor:
    """윈도우 메시지 모니터링 클래스"""

    def __init__(self, target_hwnd=None):
        self.target_hwnd = target_hwnd
        self.monitoring = False
        self.hook_id = None

    def start(self):
        """모니터링 시작"""
        self.monitoring = True
        print("\n🔍 메시지 모니터링 시작")
        print(f"   타겟 HWND: 0x{self.target_hwnd:08X}" if self.target_hwnd else "   모든 윈도우")
        print("   관심 메시지: WM_LBUTTONDOWN, WM_LBUTTONUP, WM_NOTIFY, TCM_SETCURSEL")
        print("-" * 80)

    def stop(self):
        """모니터링 중지"""
        self.monitoring = False
        print("-" * 80)
        print("🛑 메시지 모니터링 중지\n")

    def log_sent_message(self, hwnd, msg, wparam, lparam):
        """SendMessage 전송 로그"""
        if self.monitoring:
            log_message(hwnd, msg, wparam, lparam, "📤 SEND: ")


def monitor_with_polling(monitor, target_hwnd, duration=10):
    """
    폴링 방식으로 메시지 모니터링
    (실제 메시지 후킹은 복잡하므로 간단한 방식 사용)
    """
    # 참고: 실제 메시지 후킹은 DLL 인젝션이나 전역 훅이 필요합니다.
    # 여기서는 SendMessage 호출을 직접 로깅하는 방식을 사용합니다.
    pass


def test_tab_automation_with_monitoring():
    """탭 자동화 + 메시지 모니터링 통합 테스트"""
    print("=" * 80)
    print("탭 자동화 + 메시지 모니터링 테스트")
    print("=" * 80)

    # 1. 사원등록 윈도우 연결
    print("\n[1/3] 사원등록 윈도우 연결 중...")
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        print(f"✓ 연결 성공: HWND=0x{dlg.handle:08X}")
    except Exception as e:
        print(f"✗ 연결 실패: {e}")
        return

    # 2. 탭 컨트롤 찾기
    print("\n[2/3] 탭 컨트롤 찾기 중...")
    tab_control = None
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name().startswith("Afx:TabWnd:"):
                tab_control = ctrl
                break
        except:
            pass

    if tab_control is None:
        print("✗ 탭 컨트롤을 찾을 수 없습니다")
        return

    tab_hwnd = tab_control.handle
    print(f"✓ 탭 컨트롤 발견: HWND=0x{tab_hwnd:08X}")
    print(f"  클래스: {tab_control.class_name()}")

    # 3. 모니터링 시작 + 탭 선택
    print("\n[3/3] 탭 자동화 실행 + 메시지 모니터링")

    monitor = MessageMonitor(target_hwnd=tab_hwnd)
    monitor.start()

    # 탭 위치
    tab_positions = {
        "기본사항": (50, 15),
        "부양가족정보": (150, 15),
        "소득자료": (250, 15),
    }

    # 테스트할 탭들
    test_tabs = ["부양가족정보", "소득자료", "기본사항"]

    for tab_name in test_tabs:
        print(f"\n▶ '{tab_name}' 탭 선택...")
        x, y = tab_positions[tab_name]
        lparam = win32api.MAKELONG(x, y)

        # WM_LBUTTONDOWN
        print(f"  → WM_LBUTTONDOWN 전송 (x={x}, y={y})")
        monitor.log_sent_message(tab_hwnd, win32con.WM_LBUTTONDOWN,
                                win32con.MK_LBUTTON, lparam)
        result = win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONDOWN,
                                      win32con.MK_LBUTTON, lparam)
        print(f"  ← 반환값: {result}")

        time.sleep(0.1)

        # WM_LBUTTONUP
        print(f"  → WM_LBUTTONUP 전송")
        monitor.log_sent_message(tab_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        result = win32api.SendMessage(tab_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        print(f"  ← 반환값: {result}")

        time.sleep(1.5)

    monitor.stop()

    # 4. 결과 요약
    print("\n" + "=" * 80)
    print("📊 메시지 요약")
    print("=" * 80)
    print(f"총 {len(captured_messages)}개 메시지 캡처됨")

    if captured_messages:
        print("\n캡처된 메시지:")
        for i, msg in enumerate(captured_messages, 1):
            x, y = decode_lparam_coords(msg['lparam']) if msg['msg'] in [
                win32con.WM_LBUTTONDOWN, win32con.WM_LBUTTONUP
            ] else (None, None)

            coord_str = f" (x={x}, y={y})" if x is not None else ""
            print(f"  {i}. [{msg['timestamp']}] {msg['msg_name']}{coord_str}")

    print("\n💡 참고:")
    print("  - 이 스크립트는 SendMessage 호출을 직접 로깅합니다")
    print("  - 시스템 내부 메시지 (WM_NOTIFY 등)는 캡처되지 않습니다")
    print("  - 완전한 메시지 후킹은 DLL 인젝션이 필요합니다")
    print("  - 하지만 우리가 보낸 메시지는 모두 확인 가능합니다!")
    print("=" * 80)


if __name__ == "__main__":
    test_tab_automation_with_monitoring()
