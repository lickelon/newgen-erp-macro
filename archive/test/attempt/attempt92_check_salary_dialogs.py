"""
시도 92: 급여자료입력 프로그램의 다이얼로그 확인
"""
from pywinauto import Application

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 92: 급여자료입력 프로그램 다이얼로그 확인")
    print("="*60)

    try:
        # 급여자료입력 창에 연결
        print("\n[1단계] 급여자료입력 프로그램 연결")
        print("-" * 60)

        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        print(f"✓ 연결 성공")
        print(f"  제목: {main_window.window_text()}")
        print(f"  클래스: {main_window.class_name()}")
        print(f"  HWND: 0x{main_window.handle:08X}")

        # 모든 자식 창 확인
        print("\n[2단계] 자식 창/다이얼로그 확인")
        print("-" * 60)

        children = main_window.children()
        print(f"직계 자식: {len(children)}개")

        for i, child in enumerate(children):
            try:
                title = child.window_text()
                class_name = child.class_name()
                hwnd = child.handle

                print(f"\n  자식 #{i+1}")
                print(f"    HWND: 0x{hwnd:08X}")
                print(f"    클래스: {class_name}")
                if title:
                    print(f"    제목: '{title}'")
                    if "분납" in title:
                        print(f"    >>> ⭐⭐⭐ 분납 관련 발견!")
            except:
                pass

        # 모든 하위 컨트롤 확인
        print("\n[3단계] 모든 하위 컨트롤 확인 (descendants)")
        print("-" * 60)

        all_controls = main_window.descendants()
        print(f"전체 하위 컨트롤: {len(all_controls)}개")

        # 다이얼로그 클래스 찾기
        dialogs = []
        for ctrl in all_controls:
            try:
                class_name = ctrl.class_name()
                if class_name == "#32770":  # Dialog 클래스
                    dialogs.append(ctrl)
            except:
                pass

        print(f"\n다이얼로그(#32770) 발견: {len(dialogs)}개")

        found_installment = False
        for i, dialog in enumerate(dialogs):
            try:
                title = dialog.window_text()
                hwnd = dialog.handle

                print(f"\n  다이얼로그 #{i+1}")
                print(f"    HWND: 0x{hwnd:08X}")
                if title:
                    print(f"    제목: '{title}'")
                    if "분납" in title:
                        print(f"    >>> ⭐⭐⭐ 분납 관련 다이얼로그 발견!")
                        found_installment = True
                else:
                    print(f"    제목: (없음)")
            except:
                pass

        # 텍스트에 "분납" 포함된 컨트롤 찾기
        print("\n[4단계] '분납' 텍스트 포함 컨트롤 찾기")
        print("-" * 60)

        for ctrl in all_controls:
            try:
                text = ctrl.window_text()
                if text and "분납" in text:
                    print(f"\n  발견!")
                    print(f"    클래스: {ctrl.class_name()}")
                    print(f"    HWND: 0x{ctrl.handle:08X}")
                    print(f"    텍스트: '{text}'")
                    found_installment = True
            except:
                pass

        if found_installment:
            return {"success": True, "message": "분납 관련 컨트롤을 찾았습니다"}
        else:
            return {"success": False, "message": "분납 관련 컨트롤을 찾지 못했습니다. 분납적용 창/다이얼로그가 열려있는지 확인하세요."}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
