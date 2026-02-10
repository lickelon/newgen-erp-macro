"""
분납적용 다이얼로그의 모든 Edit 컨트롤 찾기
"""
import sys
import win32process
import win32gui
import win32con
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


def find_all_edit_controls(dialog):
    """다이얼로그 내 모든 Edit 컨트롤 찾기"""
    edit_controls = []

    def find_edits_recursive(parent, depth=0):
        for child in parent.children():
            try:
                class_name = child.class_name()
                text = child.window_text()
                rect = child.rectangle()

                # Edit 또는 입력 가능한 컨트롤
                if class_name in ["Edit", "ThunderRT6TextBox", "fpUSpread80"]:
                    edit_controls.append({
                        "class": class_name,
                        "text": text if text else "(empty)",
                        "hwnd": child.handle,
                        "rect": rect,
                        "depth": depth
                    })

                # 재귀적으로 자식 탐색
                if child.children():
                    find_edits_recursive(child, depth + 1)

            except:
                pass

    find_edits_recursive(dialog)
    return edit_controls


def main():
    print("="*70)
    print("분납적용 다이얼로그의 Edit 컨트롤 찾기")
    print("="*70)

    try:
        # 급여자료입력 연결
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n급여자료입력 연결 (PID: {process_id})")

        # 분납적용 다이얼로그 찾기
        installment_dlg, dlg_hwnd = find_installment_dialog(app, process_id)

        if not installment_dlg:
            print("❌ 분납적용 다이얼로그를 찾지 못했습니다!")
            return

        print(f"✓ 분납적용 다이얼로그: 0x{dlg_hwnd:08X}")

        # Edit 컨트롤 찾기
        print("\n모든 Edit 컨트롤 찾기:")
        edit_controls = find_all_edit_controls(installment_dlg)

        if not edit_controls:
            print("  Edit 컨트롤을 찾지 못했습니다!")
        else:
            print(f"\n총 {len(edit_controls)}개의 입력 가능한 컨트롤 발견:")
            for idx, ctrl in enumerate(edit_controls):
                print(f"\n[{idx+1}] {ctrl['class']}")
                print(f"  HWND: 0x{ctrl['hwnd']:08X}")
                print(f"  Text: {ctrl['text']}")
                print(f"  위치: left={ctrl['rect'].left}, top={ctrl['rect'].top}, right={ctrl['rect'].right}, bottom={ctrl['rect'].bottom}")
                print(f"  Depth: {ctrl['depth']}")

    except Exception as e:
        import traceback
        print(f"\n❌ 오류: {e}")
        print(traceback.format_exc())


if __name__ == "__main__":
    main()
