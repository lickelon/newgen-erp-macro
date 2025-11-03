"""
Attempt 14: 우측 패널의 실제 입력 필드 찾기

왼쪽 = 사원 목록 (fpUSpread80)
우측 = 상세 정보 입력 폼
→ 우측 폼의 Edit 컨트롤이나 다른 fpUSpread 찾기
"""
import time
import win32api
import win32con
import win32gui


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 14: 우측 패널의 실제 입력 필드 찾기")
    print("=" * 60)

    try:
        capture_func("attempt14_00_initial.png")

        # 기본사항 탭
        print("\n[1/5] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt14_01_basic_tab.png")

        main_rect = dlg.rectangle()
        main_hwnd = dlg.handle

        print(f"\n[2/5] 메인 윈도우: L={main_rect.left} T={main_rect.top}")
        print(f"                 R={main_rect.right} B={main_rect.bottom}")
        print(f"                 중심: {main_rect.mid_point()}")

        # 우측 패널 영역 추정 (메인 윈도우의 오른쪽 절반)
        mid_x = main_rect.mid_point()[0]

        # 우측 영역의 컨트롤 찾기
        print(f"\n[3/5] 우측 영역 (x > {mid_x}) 컨트롤 스캔...")

        right_controls = []
        for ctrl in dlg.descendants():
            try:
                rect = ctrl.rectangle()
                class_name = ctrl.class_name()

                # 우측 영역에 있는 컨트롤
                if rect.left > mid_x:
                    # Edit, fpUSpread, 또는 입력 가능한 컨트롤
                    if any(keyword in class_name for keyword in ["Edit", "fpUSpread", "TextBox", "Input"]):
                        right_controls.append({
                            "ctrl": ctrl,
                            "class": class_name,
                            "rect": rect,
                            "hwnd": ctrl.handle,
                            "visible": ctrl.is_visible(),
                            "enabled": ctrl.is_enabled()
                        })
            except:
                pass

        print(f"  우측 영역 컨트롤: {len(right_controls)}개")
        for idx, info in enumerate(right_controls):
            rect = info["rect"]
            print(f"  [{idx}] {info['class']}")
            print(f"       HWND=0x{info['hwnd']:08X}")
            print(f"       위치=({rect.left},{rect.top}) 크기={rect.width()}x{rect.height()}")
            print(f"       visible={info['visible']} enabled={info['enabled']}")
            try:
                text = info["ctrl"].window_text()
                if text:
                    print(f"       text=\"{text}\"")
            except:
                pass

        # Static 라벨 찾기 (우측)
        print(f"\n[4/5] 우측 영역의 Static 라벨...")
        labels = []
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "Static":
                    rect = ctrl.rectangle()
                    if rect.left > mid_x:
                        text = ctrl.window_text()
                        if text and any(kw in text for kw in ["사번", "성명", "주민", "이름"]):
                            labels.append({
                                "text": text,
                                "rect": rect,
                                "ctrl": ctrl
                            })
            except:
                pass

        print(f"  관련 라벨: {len(labels)}개")
        for label in labels:
            print(f"  • \"{label['text']}\" at ({label['rect'].left},{label['rect'].top})")

        # 입력 시도
        print(f"\n[5/5] 우측 컨트롤에 입력 시도...")

        test_data = [
            ("필드1", "2025999"),
            ("필드2", "홍길동"),
            ("필드3", "900101-1234567"),
        ]

        # fpUSpread가 있으면 클릭 + WM_CHAR
        fpuspread_in_right = [c for c in right_controls if "fpUSpread" in c["class"]]
        edit_in_right = [c for c in right_controls if c["class"] == "Edit" and c["visible"]]

        if fpuspread_in_right:
            print(f"\n  fpUSpread 컨트롤 {len(fpuspread_in_right)}개 발견")
            spread = fpuspread_in_right[0]
            hwnd = spread["hwnd"]
            rect = spread["rect"]

            # 상대 좌표 (오프셋)
            offsets = [
                (100, 50),
                (100, 100),
                (100, 150),
            ]

            for idx, ((offset_x, offset_y), (label, value)) in enumerate(zip(offsets, test_data)):
                print(f"\n  [{idx}] {label} = \"{value}\" at ({offset_x},{offset_y})")

                try:
                    lparam = win32api.MAKELONG(offset_x, offset_y)
                    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                    time.sleep(0.05)
                    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                    time.sleep(0.2)

                    for char in value:
                        win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                        time.sleep(0.02)

                    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                    time.sleep(0.02)
                    win32api.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                    time.sleep(0.3)

                    print(f"       ✓ 입력 완료")

                except Exception as e:
                    print(f"       ✗ 오류: {e}")

        elif edit_in_right:
            print(f"\n  Edit 컨트롤 {len(edit_in_right)}개 발견")
            for idx, info in enumerate(edit_in_right[:3]):
                if idx >= len(test_data):
                    break

                label, value = test_data[idx]
                hwnd = info["hwnd"]

                print(f"  [{idx}] {label} = \"{value}\" HWND=0x{hwnd:08X}")

                try:
                    win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, value)
                    time.sleep(0.1)

                    # EN_CHANGE 알림
                    parent = win32gui.GetParent(hwnd)
                    if parent:
                        ctrl_id = win32api.GetWindowLong(hwnd, win32con.GWL_ID)
                        wparam = (0x0300 << 16) | ctrl_id  # EN_CHANGE
                        win32api.SendMessage(parent, 0x0111, wparam, hwnd)  # WM_COMMAND
                    time.sleep(0.1)

                    result = win32gui.GetWindowText(hwnd)
                    print(f"       결과: \"{result}\" {'✅' if result == value else '❌'}")

                except Exception as e:
                    print(f"       ✗ 오류: {e}")

        else:
            print("  ⚠️  입력 가능한 컨트롤을 찾지 못함")

        time.sleep(1)
        capture_func("attempt14_02_after_input.png")

        return {"success": True, "message": "우측 패널 입력 시도 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}


if __name__ == "__main__":
    from pywinauto import application
    import sys
    import os

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    result = run(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
