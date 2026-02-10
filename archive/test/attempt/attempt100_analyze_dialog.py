"""
시도 100: 제목 없는 다이얼로그 분석
"""
from pywinauto import Application
import win32process
import win32gui

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 스크린샷 함수

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 100: 제목 없는 다이얼로그 분석")
    print("="*60)

    try:
        # 급여자료입력 연결
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        # 프로세스 ID
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        # 같은 프로세스의 #32770 창 찾기
        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    try:
                        class_name = win32gui.GetClassName(hwnd)
                        if class_name == "#32770":
                            title = win32gui.GetWindowText(hwnd)
                            results.append((hwnd, title))
                    except:
                        pass
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        print(f"\n같은 프로세스의 #32770 다이얼로그: {len(found_dialogs)}개")

        for hwnd, title in found_dialogs:
            print(f"\n{'='*60}")
            print(f"다이얼로그 HWND: 0x{hwnd:08X}")
            print(f"제목: '{title}'" if title else "제목: (없음)")
            print('='*60)

            # pywinauto로 연결
            try:
                dialog_win = app.window(handle=hwnd)

                # 자식 컨트롤 확인
                children = dialog_win.children()
                print(f"\n자식 컨트롤: {len(children)}개")

                # 컨트롤 타입별 분류
                control_types = {}
                for child in children:
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

                # 버튼과 Static 텍스트 확인
                print(f"\n주요 컨트롤:")
                for child in children[:20]:  # 처음 20개만
                    try:
                        cn = child.class_name()
                        text = child.window_text()
                        if text and cn in ["Button", "Static"]:
                            print(f"  [{cn}] '{text}'")
                            if "분납" in text:
                                print(f"    >>> ⭐⭐⭐ 분납 관련!")
                    except:
                        pass

            except Exception as e:
                print(f"분석 실패: {e}")

        return {"success": True, "message": f"{len(found_dialogs)}개 다이얼로그 분석 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
