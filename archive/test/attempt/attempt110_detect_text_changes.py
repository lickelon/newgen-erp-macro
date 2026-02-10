"""
시도 110: 텍스트 변화 감지

행 클릭 전후로 모든 윈도우/컨트롤의 텍스트 변화를 감지하여
데이터가 어디에 표시되는지 찾기
"""
import win32process
import win32gui
import win32con
import win32api
import time
from pywinauto import Application
from ctypes import create_unicode_buffer


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 110: 텍스트 변화 감지")
    print("="*60)

    try:
        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[단계] 급여자료입력 연결 (PID: {process_id})")

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

        installment_dlg = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 2. 스프레드 찾기
        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        left_hwnd = left_spread.handle
        print(f"\n왼쪽 스프레드: 0x{left_hwnd:08X}")

        # === 모든 윈도우/컨트롤 텍스트 수집 함수 ===
        def collect_all_texts():
            """모든 윈도우와 컨트롤의 텍스트 수집"""
            texts = {}

            # 같은 프로세스의 모든 윈도우
            all_windows = []

            def enum_proc(hwnd, lParam):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    all_windows.append(hwnd)
                return True

            win32gui.EnumWindows(enum_proc, None)

            for hwnd in all_windows:
                try:
                    # 윈도우 텍스트
                    text = win32gui.GetWindowText(hwnd)
                    if text:
                        cn = win32gui.GetClassName(hwnd)
                        key = f"0x{hwnd:08X}_{cn}"
                        texts[key] = text

                    # 자식 윈도우들
                    child_windows = []

                    def enum_child(child_hwnd, lParam):
                        child_windows.append(child_hwnd)
                        return True

                    win32gui.EnumChildWindows(hwnd, enum_child, None)

                    for child_hwnd in child_windows:
                        try:
                            child_text = win32gui.GetWindowText(child_hwnd)
                            if child_text:
                                child_cn = win32gui.GetClassName(child_hwnd)
                                child_key = f"0x{child_hwnd:08X}_{child_cn}"
                                texts[child_key] = child_text

                            # WM_GETTEXT도 시도 (더 큰 버퍼)
                            buffer = create_unicode_buffer(1024)
                            length = win32gui.SendMessage(child_hwnd, win32con.WM_GETTEXT, 1024, buffer)
                            if length > 0:
                                wm_text = buffer.value
                                if wm_text and wm_text != child_text:
                                    child_key2 = f"0x{child_hwnd:08X}_{child_cn}_WM"
                                    texts[child_key2] = wm_text
                        except:
                            pass
                except:
                    pass

            return texts

        # === 방법 1: 첫 번째 행 클릭 전후 비교 ===
        print("\n" + "="*60)
        print("[방법 1] 첫 번째 행 클릭 전후 텍스트 변화")
        print("="*60)

        print("클릭 전 텍스트 수집 중...")
        before_texts = collect_all_texts()
        print(f"✓ 클릭 전: {len(before_texts)}개 텍스트")

        # 첫 번째 행 클릭
        lParam = win32api.MAKELONG(50, 30)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.05)
        win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
        time.sleep(0.5)

        print("클릭 후 텍스트 수집 중...")
        after_texts = collect_all_texts()
        print(f"✓ 클릭 후: {len(after_texts)}개 텍스트")

        # 차이 분석
        print("\n변화 감지:")

        # 새로 생긴 텍스트
        new_keys = set(after_texts.keys()) - set(before_texts.keys())
        if new_keys:
            print(f"\n✓ 새로 생긴 텍스트 {len(new_keys)}개:")
            for key in list(new_keys)[:10]:  # 처음 10개만
                print(f"  [{key}] = '{after_texts[key]}'")

                # 사원코드처럼 보이는 패턴 찾기
                text = after_texts[key]
                if text and (text.isdigit() or (len(text) < 20 and not any(x in text for x in ['분납', '환급', '취소']))):
                    print(f"\n✓✓✓ 사원코드 후보 발견: '{text}'")
                    return {"success": True, "message": f"새 텍스트에서 발견: '{text}' (키: {key})"}

        # 변경된 텍스트
        changed_keys = []
        for key in before_texts:
            if key in after_texts and before_texts[key] != after_texts[key]:
                changed_keys.append(key)

        if changed_keys:
            print(f"\n✓ 변경된 텍스트 {len(changed_keys)}개:")
            for key in changed_keys[:10]:
                print(f"  [{key}]")
                print(f"    전: '{before_texts[key]}'")
                print(f"    후: '{after_texts[key]}'")

                # 의미있는 변화인지 확인
                before = before_texts[key]
                after = after_texts[key]
                if after and after != before and len(after) < 50:
                    print(f"\n✓✓✓ 텍스트 변화 감지: '{before}' → '{after}'")
                    return {"success": True, "message": f"텍스트 변화: '{after}' (키: {key})"}

        if not new_keys and not changed_keys:
            print("✗ 텍스트 변화 없음")

        # === 방법 2: 여러 행 순회하며 패턴 찾기 ===
        print("\n" + "="*60)
        print("[방법 2] 여러 행 클릭하며 패턴 찾기")
        print("="*60)

        # Home으로 시작
        win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_HOME, 0)
        win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_HOME, 0)
        time.sleep(0.3)

        row_data = []

        for row in range(3):
            if row > 0:
                # Down 키
                win32api.SendMessage(left_hwnd, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
                win32api.SendMessage(left_hwnd, win32con.WM_KEYUP, win32con.VK_DOWN, 0)
                time.sleep(0.3)

            # 현재 텍스트 수집
            current_texts = collect_all_texts()

            # 짧은 텍스트만 추출 (사원코드 후보)
            short_texts = []
            for key, text in current_texts.items():
                if text and 1 <= len(text) <= 20:
                    # 일반 안내 문구 제외
                    if not any(x in text for x in ['분납', '환급', '취소', '계산', '적용', '인쇄', 'Tab', 'Esc', 'F9']):
                        short_texts.append(text)

            row_data.append(set(short_texts))
            print(f"  행 {row}: {len(short_texts)}개 짧은 텍스트")

        # 행마다 고유한 텍스트 찾기
        print("\n행별 고유 텍스트 분석:")
        for i in range(len(row_data)):
            # 다른 행에는 없고 이 행에만 있는 텍스트
            unique = row_data[i]
            for j in range(len(row_data)):
                if i != j:
                    unique = unique - row_data[j]

            if unique:
                print(f"  행 {i} 고유: {list(unique)[:5]}")

        # === 방법 3: 오른쪽 스프레드의 변화 확인 ===
        if len(spreads) > 1:
            print("\n" + "="*60)
            print("[방법 3] 오른쪽 스프레드 텍스트 변화")
            print("="*60)

            right_spread = spreads[1]
            right_hwnd = right_spread.handle

            # 왼쪽에서 행 클릭
            lParam = win32api.MAKELONG(50, 50)
            win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
            time.sleep(0.05)
            win32api.SendMessage(left_hwnd, win32con.WM_LBUTTONUP, 0, lParam)
            time.sleep(0.5)

            # 오른쪽 스프레드에서 복사 시도
            win32api.SendMessage(right_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYDOWN, ord('C'), 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYUP, ord('C'), 0)
            win32api.SendMessage(right_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
            time.sleep(0.3)

            import pyperclip
            value = pyperclip.paste()
            if value:
                print(f"✓✓✓ 오른쪽 스프레드에서 복사 성공: '{value}'")
                return {"success": True, "message": f"오른쪽 스프레드 복사: '{value}'"}
            else:
                print("✗ 오른쪽 스프레드도 복사 안 됨")

        return {
            "success": False,
            "message": "텍스트 변화 감지 실패 - 데이터가 윈도우 텍스트로 노출되지 않음"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
