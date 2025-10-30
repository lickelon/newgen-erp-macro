"""
Attempt 10: 보이는(visible) 컨트롤에만 입력

이전 시도는 보이지 않는 컨트롤에 입력했을 가능성
→ visible=True인 컨트롤만 찾아서 입력
"""
import time
import win32api
import win32con
import win32gui


def run(dlg, capture_func):
    print("\n" + "=" * 60)
    print("Attempt 10: 보이는(visible) 컨트롤에만 입력")
    print("=" * 60)

    try:
        capture_func("attempt10_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt10_01_basic_tab.png")

        # 모든 Edit 컨트롤 찾기 (클래스명 종류별)
        print("\n[2/4] Edit 컨트롤 분석...")
        edit_types = {}
        for ctrl in dlg.descendants():
            try:
                class_name = ctrl.class_name()
                if "Edit" in class_name or "edit" in class_name:
                    visible = ctrl.is_visible()
                    enabled = ctrl.is_enabled()
                    rect = ctrl.rectangle()
                    text = ctrl.window_text()

                    key = f"{class_name}|visible={visible}|enabled={enabled}"
                    if key not in edit_types:
                        edit_types[key] = []
                    edit_types[key].append({
                        "ctrl": ctrl,
                        "text": text,
                        "rect": rect,
                        "hwnd": f"0x{ctrl.handle:08X}"
                    })
            except:
                pass

        print(f"  발견된 Edit 타입: {len(edit_types)}개")
        for key, ctrls in edit_types.items():
            print(f"  • {key}: {len(ctrls)}개")

        # visible=True, enabled=True인 일반 Edit 컨트롤만
        print("\n[3/4] 보이는 Edit 컨트롤 필터링...")
        visible_edits = []
        for ctrl in dlg.descendants():
            try:
                class_name = ctrl.class_name()
                if class_name == "Edit":  # 일반 Edit만
                    if ctrl.is_visible() and ctrl.is_enabled():
                        rect = ctrl.rectangle()
                        # 크기가 있는 것만 (0x0 제외)
                        if rect.width() > 0 and rect.height() > 0:
                            visible_edits.append(ctrl)
            except:
                pass

        print(f"  보이는 Edit 컨트롤: {len(visible_edits)}개")
        for idx, ctrl in enumerate(visible_edits):
            rect = ctrl.rectangle()
            text = ctrl.window_text()
            print(f"    [{idx}] HWND=0x{ctrl.handle:08X} text=\"{text}\"")
            print(f"        위치: ({rect.left},{rect.top}) 크기: {rect.width()}x{rect.height()}")

        if not visible_edits:
            return {"success": False, "message": "보이는 Edit 컨트롤 없음"}

        # 테스트 데이터
        test_data = [
            ("필드1", "TEST001"),
            ("필드2", "900101-1234567"),
            ("필드3", "테스트"),
        ]

        print(f"\n[4/4] 데이터 입력 시도...")
        EN_CHANGE = 0x0300
        WM_COMMAND = 0x0111

        for idx, ctrl in enumerate(visible_edits[:3]):  # 처음 3개만
            if idx >= len(test_data):
                break

            label, value = test_data[idx]
            hwnd = ctrl.handle

            print(f"\n  [{idx}] {label} = \"{value}\"")
            print(f"      HWND=0x{hwnd:08X}")

            try:
                # 클릭 시뮬레이션 (포커스)
                rect = ctrl.rectangle()
                x = rect.width() // 2
                y = rect.height() // 2
                lparam = win32api.MAKELONG(x, y)

                print(f"      1. 클릭 시뮬레이션")
                win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                time.sleep(0.05)
                win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                time.sleep(0.1)

                # 전체 선택
                print(f"      2. 전체 선택 (EM_SETSEL)")
                EM_SETSEL = 0x00B1
                win32api.SendMessage(hwnd, EM_SETSEL, 0, -1)
                time.sleep(0.05)

                # 텍스트 설정
                print(f"      3. WM_SETTEXT")
                win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, value)
                time.sleep(0.1)

                # EN_CHANGE
                print(f"      4. EN_CHANGE 알림")
                parent = win32gui.GetParent(hwnd)
                if parent:
                    ctrl_id = win32api.GetWindowLong(hwnd, win32con.GWL_ID)
                    wparam = (EN_CHANGE << 16) | ctrl_id
                    win32api.SendMessage(parent, WM_COMMAND, wparam, hwnd)
                time.sleep(0.1)

                # 결과 확인
                result = ctrl.window_text()
                print(f"      결과: \"{result}\" {'✅' if result == value else '❌'}")

            except Exception as e:
                print(f"      ✗ 오류: {e}")

        time.sleep(1)
        capture_func("attempt10_02_after_input.png")

        return {"success": True, "message": "보이는 Edit 컨트롤 입력 시도 완료"}

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
