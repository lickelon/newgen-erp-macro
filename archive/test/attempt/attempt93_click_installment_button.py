"""
시도 93: 분납적용 버튼 클릭 및 다이얼로그 확인
"""
import time
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
    print("시도 93: 분납적용 버튼 클릭 및 다이얼로그 확인")
    print("="*60)

    try:
        # 급여자료입력 창에 연결
        print("\n[1단계] 급여자료입력 프로그램 연결")
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        print(f"✓ 연결 성공: 0x{main_window.handle:08X}")

        # 초기 상태 캡처
        capture_func("attempt93_00_initial.png")

        # 분납적용 버튼 찾기
        print("\n[2단계] 분납적용 버튼 찾기")

        installment_button = None
        for child in main_window.children():
            try:
                if child.class_name() == "Button" and "분납적용" in child.window_text():
                    installment_button = child
                    print(f"✓ 버튼 발견: 0x{child.handle:08X}")
                    break
            except:
                pass

        if not installment_button:
            return {"success": False, "message": "분납적용 버튼을 찾지 못했습니다"}

        # 버튼 클릭
        print("\n[3단계] 버튼 클릭")
        installment_button.click_input()
        time.sleep(1.0)

        capture_func("attempt93_01_after_click.png")

        # 다이얼로그 확인
        print("\n[4단계] 분납적용 다이얼로그 찾기")

        # 모든 다이얼로그 찾기
        all_controls = main_window.descendants()

        dialogs = []
        for ctrl in all_controls:
            try:
                if ctrl.class_name() == "#32770":
                    title = ctrl.window_text()
                    if title and "분납" in title:
                        dialogs.append(ctrl)
                        print(f"✓ 분납 다이얼로그 발견!")
                        print(f"  HWND: 0x{ctrl.handle:08X}")
                        print(f"  제목: '{title}'")
            except:
                pass

        if not dialogs:
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

            return {"success": False, "message": "분납 다이얼로그를 찾지 못했습니다"}

        # 다이얼로그 구조 확인
        print("\n[5단계] 다이얼로그 구조 확인")
        installment_dlg = dialogs[0]

        print(f"\n기본 정보:")
        print(f"  제목: {installment_dlg.window_text()}")
        print(f"  클래스: {installment_dlg.class_name()}")
        print(f"  HWND: 0x{installment_dlg.handle:08X}")

        # 다이얼로그의 컨트롤 확인
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

        capture_func("attempt93_02_dialog_opened.png")

        return {"success": True, "message": f"분납적용 다이얼로그 열기 성공 (0x{installment_dlg.handle:08X})"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
