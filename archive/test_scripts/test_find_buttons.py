"""
분납적용 다이얼로그의 모든 버튼 찾기
"""
import sys
import win32process
import win32gui
from pywinauto import Application

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def find_installment_dialog(app, process_id):
    """분납적용 다이얼로그 찾기"""
    found_dialogs = []

    def enum_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == process_id:
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "#32770":
                    title = win32gui.GetWindowText(hwnd)
                    results.append((hwnd, title))
        return True

    win32gui.EnumWindows(enum_callback, found_dialogs)

    for hwnd, title in found_dialogs:
        if not title:
            dialog = app.window(handle=hwnd)
            for child in dialog.children():
                try:
                    text = child.window_text()
                    if "분납적용" in text:
                        return dialog, hwnd
                except:
                    pass

    return None, None


def main():
    print("="*70)
    print("분납적용 다이얼로그의 버튼 찾기")
    print("="*70)

    try:
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n급여자료입력 연결 (PID: {process_id})")

        installment_dlg, dlg_hwnd = find_installment_dialog(app, process_id)

        if not installment_dlg:
            print("❌ 분납적용 다이얼로그를 찾지 못했습니다!")
            return

        print(f"✓ 분납적용 다이얼로그: 0x{dlg_hwnd:08X}\n")

        # 모든 버튼 찾기
        print("모든 버튼:")
        buttons = []
        for child in installment_dlg.children():
            try:
                class_name = child.class_name()
                if "Button" in class_name:
                    text = child.window_text()
                    hwnd = child.handle
                    rect = child.rectangle()
                    buttons.append({
                        "class": class_name,
                        "text": text,
                        "hwnd": hwnd,
                        "rect": rect
                    })
            except:
                pass

        if buttons:
            for idx, btn in enumerate(buttons):
                print(f"\n[{idx+1}] {btn['text']}")
                print(f"  Class: {btn['class']}")
                print(f"  HWND: 0x{btn['hwnd']:08X}")
                print(f"  위치: left={btn['rect'].left}, top={btn['rect'].top}")
        else:
            print("  버튼을 찾지 못했습니다!")

    except Exception as e:
        import traceback
        print(f"\n❌ 오류: {e}")
        print(traceback.format_exc())


if __name__ == "__main__":
    main()
