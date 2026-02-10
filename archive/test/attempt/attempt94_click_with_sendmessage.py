"""
시도 94: SendMessage로 분납적용 버튼 클릭
"""
import time
import win32api
import win32con
from pywinauto import Application

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 스크린샷 함수

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 94: SendMessage로 분납적용 버튼 클릭")
    print("="*60)

    try:
        # 급여자료입력 창에 연결
        print("\n[1단계] 급여자료입력 프로그램 연결")
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        print(f"✓ 연결 성공: 0x{main_window.handle:08X}")

        # 초기 상태 캡처
        capture_func("attempt94_00_initial.png")

        # 분납적용 버튼 찾기
        print("\n[2단계] 분납적용 버튼 찾기")

        installment_button = None
        for child in main_window.children():
            try:
                if child.class_name() == "Button" and "분납적용" in child.window_text():
                    installment_button = child
                    button_hwnd = child.handle
                    print(f"✓ 버튼 발견: 0x{button_hwnd:08X}")
                    break
            except:
                pass

        if not installment_button:
            return {"success": False, "message": "분납적용 버튼을 찾지 못했습니다"}

        # SendMessage로 버튼 클릭 (BM_CLICK)
        print("\n[3단계] SendMessage로 버튼 클릭")
        print(f"  BM_CLICK 메시지 전송 (0xF5)")

        win32api.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
        time.sleep(1.5)

        capture_func("attempt94_01_after_click.png")

        # 다이얼로그 확인
        print("\n[4단계] 분납적용 다이얼로그 찾기")

        # 모든 다이얼로그 찾기
        all_controls = main_window.descendants()

        found_dialogs = []
        for ctrl in all_controls:
            try:
                if ctrl.class_name() == "#32770":
                    title = ctrl.window_text()
                    if title and "분납" in title:
                        found_dialogs.append(ctrl)
                        print(f"✓ 분납 다이얼로그 발견!")
                        print(f"  HWND: 0x{ctrl.handle:08X}")
                        print(f"  제목: '{title}'")
            except:
                pass

        if not found_dialogs:
            print("⚠️  분납 관련 다이얼로그를 찾지 못했습니다")
            print("    모든 #32770 다이얼로그 확인:")

            for ctrl in all_controls:
                try:
                    if ctrl.class_name() == "#32770":
                        title = ctrl.window_text()
                        if title:
                            print(f"    - '{title}' (0x{ctrl.handle:08X})")
                except:
                    pass

            # 독립 창으로 열렸을 가능성 확인
            print("\n    독립 창으로 열렸는지 확인 중...")
            from pywinauto import Desktop
            desktop = Desktop(backend="win32")
            for win in desktop.windows():
                try:
                    title = win.window_text()
                    if title and "분납" in title:
                        print(f"    ✓ 독립 창 발견: '{title}' (0x{win.handle:08X})")
                        return {"success": True, "message": f"분납적용 독립 창 열림 (0x{win.handle:08X})"}
                except:
                    pass

            return {"success": False, "message": "분납 다이얼로그를 찾지 못했습니다"}

        # 다이얼로그 구조 확인
        print("\n[5단계] 다이얼로그 구조 확인")
        installment_dlg = found_dialogs[0]

        print(f"\n기본 정보:")
        print(f"  제목: {installment_dlg.window_text()}")
        print(f"  클래스: {installment_dlg.class_name()}")
        print(f"  HWND: 0x{installment_dlg.handle:08X}")

        # 컨트롤 타입별 분류
        dlg_children = installment_dlg.children()
        print(f"\n자식 컨트롤: {len(dlg_children)}개")

        control_types = {}
        for child in dlg_children:
            try:
                cn = child.class_name()
                if cn not in control_types:
                    control_types[cn] = 0
                control_types[cn] += 1
            except:
                pass

        print(f"\n컨트롤 타입별:")
        for cn, count in sorted(control_types.items()):
            print(f"  {cn}: {count}개")

        capture_func("attempt94_02_dialog_opened.png")

        return {"success": True, "message": f"분납적용 다이얼로그 열기 성공 (0x{installment_dlg.handle:08X})"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
