"""
시도 96: 분납적용 다이얼로그 정확히 찾기
"""
from pywinauto import Application, Desktop

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 스크린샷 함수

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 96: 분납적용 다이얼로그 정확히 찾기")
    print("="*60)

    try:
        # 방법 1: 급여자료입력의 자식으로 찾기
        print("\n[방법 1] 급여자료입력 자식 다이얼로그")
        print("-" * 60)

        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        all_controls = main_window.descendants()
        found = False

        for ctrl in all_controls:
            try:
                if ctrl.class_name() == "#32770":
                    title = ctrl.window_text()
                    print(f"  다이얼로그 발견: '{title}' (0x{ctrl.handle:08X})")

                    if "분납" in title:
                        print(f"    ✓✓✓ 분납 관련 다이얼로그!")
                        found = True

                        # 상세 정보
                        print(f"\n    [상세 정보]")
                        print(f"      제목: '{title}'")
                        print(f"      클래스: {ctrl.class_name()}")
                        print(f"      HWND: 0x{ctrl.handle:08X}")

                        # 자식 컨트롤 확인
                        children = ctrl.children()
                        print(f"      자식 컨트롤: {len(children)}개")

                        return {"success": True, "message": f"분납적용 다이얼로그 발견 (0x{ctrl.handle:08X})"}
            except:
                pass

        if not found:
            print("  분납 다이얼로그 못 찾음")

        # 방법 2: Desktop의 모든 최상위 창 확인
        print("\n[방법 2] Desktop 최상위 창")
        print("-" * 60)

        desktop = Desktop(backend="win32")
        for win in desktop.windows():
            try:
                title = win.window_text()
                if title and "분납" in title:
                    print(f"  ✓ 발견: '{title}' (0x{win.handle:08X})")
                    print(f"    클래스: {win.class_name()}")
                    return {"success": True, "message": f"분납적용 창 발견 (0x{win.handle:08X})"}
            except:
                pass

        print("  분납 관련 창 못 찾음")

        # 방법 3: 모든 #32770 출력 (디버깅)
        print("\n[방법 3] 모든 #32770 다이얼로그 출력")
        print("-" * 60)

        for ctrl in all_controls:
            try:
                if ctrl.class_name() == "#32770":
                    title = ctrl.window_text()
                    print(f"  '{title}' (0x{ctrl.handle:08X})")
            except:
                pass

        return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
