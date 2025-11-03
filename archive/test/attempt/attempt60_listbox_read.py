"""
시도 60: ListBox에서 데이터 읽기

fpUSpread80 내부의 ListBox 자식 윈도우에서
LB_ 메시지를 사용하여 데이터를 읽어보기
"""
import time
import win32api
import win32con
import win32gui
from ctypes import create_unicode_buffer, c_int, windll, byref, sizeof


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 60: ListBox에서 데이터 읽기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt60_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 포커스 설정
        left_spread.set_focus()
        time.sleep(0.5)

        # 참조 값 확인 (복사로)
        import pyperclip
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (복사): '{reference_value}'")

        # ListBox 찾기
        print("\n=== ListBox 자식 찾기 ===")
        listbox_hwnd = None

        def find_listbox_callback(hwnd_child, param):
            try:
                class_name = win32gui.GetClassName(hwnd_child)
                if class_name == "ListBox":
                    param.append(hwnd_child)
                    print(f"✓ ListBox 발견: 0x{hwnd_child:08X}")
            except:
                pass
            return True

        listboxes = []
        win32gui.EnumChildWindows(hwnd, find_listbox_callback, listboxes)

        if not listboxes:
            return {"success": False, "message": "ListBox를 찾을 수 없음"}

        listbox_hwnd = listboxes[0]
        print(f"\nListBox HWND: 0x{listbox_hwnd:08X}")

        # ListBox 메시지들
        LB_GETCOUNT = 0x018B
        LB_GETCURSEL = 0x0188
        LB_GETSEL = 0x0187
        LB_GETTEXT = 0x0189
        LB_GETTEXTLEN = 0x018A
        LB_GETITEMDATA = 0x0199
        LB_GETTOPINDEX = 0x018E
        LB_GETCARETINDEX = 0x019F

        print("\n=== ListBox 정보 읽기 ===")

        # 1. 아이템 개수
        count = win32gui.SendMessage(listbox_hwnd, LB_GETCOUNT, 0, 0)
        print(f"1. 아이템 개수: {count}")

        # 2. 현재 선택된 인덱스
        cur_sel = win32gui.SendMessage(listbox_hwnd, LB_GETCURSEL, 0, 0)
        print(f"2. 현재 선택: {cur_sel}")

        # 3. Top 인덱스
        top_index = win32gui.SendMessage(listbox_hwnd, LB_GETTOPINDEX, 0, 0)
        print(f"3. Top 인덱스: {top_index}")

        # 4. Caret 인덱스
        caret_index = win32gui.SendMessage(listbox_hwnd, LB_GETCARETINDEX, 0, 0)
        print(f"4. Caret 인덱스: {caret_index}")

        print("\n=== 아이템 텍스트 읽기 ===")

        # 읽을 인덱스 결정 (현재 선택, caret, top 순서로 시도)
        indices_to_try = [cur_sel, caret_index, top_index, 0]
        indices_to_try = [i for i in indices_to_try if i >= 0 and i < count]

        found_match = False

        for idx in indices_to_try[:10]:  # 처음 10개만
            # 텍스트 길이 확인
            text_len = win32gui.SendMessage(listbox_hwnd, LB_GETTEXTLEN, idx, 0)

            if text_len > 0:
                # 텍스트 읽기
                buffer = create_unicode_buffer(text_len + 1)
                win32gui.SendMessage(listbox_hwnd, LB_GETTEXT, idx, buffer)
                text = buffer.value

                print(f"  [{idx}] (len={text_len}): '{text}'")

                # 참조 값과 비교
                if text == reference_value or reference_value in text or text in reference_value:
                    print(f"    ✓✓ 참조 값과 일치/포함!")
                    found_match = True

        # 모든 아이템 스캔 (최대 100개)
        if not found_match and count < 100:
            print(f"\n전체 {count}개 아이템 스캔 중...")
            for idx in range(count):
                text_len = win32gui.SendMessage(listbox_hwnd, LB_GETTEXTLEN, idx, 0)
                if text_len > 0:
                    buffer = create_unicode_buffer(text_len + 1)
                    win32gui.SendMessage(listbox_hwnd, LB_GETTEXT, idx, buffer)
                    text = buffer.value

                    if reference_value in text or text in reference_value:
                        print(f"  ✓ [{idx}]: '{text}' - 일치!")
                        found_match = True
                        break

        capture_func("attempt60_01_complete.png")

        if found_match:
            return {
                "success": True,
                "message": f"ListBox에서 데이터 읽기 성공! 인덱스={idx}, 값='{text}'"
            }
        else:
            return {
                "success": False,
                "message": f"ListBox에 {count}개 아이템 있지만 일치하는 값 없음"
            }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
