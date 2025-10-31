"""
시도 72: 탭 구조 탐색 및 기본사항 탭 클릭

모든 컨트롤을 출력하여 구조 파악
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 72: 탭 구조 탐색")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt72_00_initial.png")

        print("\n=== 1. 모든 컨트롤 트리 출력 ===")

        def print_control_tree(hwnd, indent=0):
            """컨트롤 트리를 재귀적으로 출력"""
            try:
                class_name = win32gui.GetClassName(hwnd)
                text = win32gui.GetWindowText(hwnd)
                rect = win32gui.GetWindowRect(hwnd)

                # 출력
                prefix = "  " * indent
                if text:
                    print(f"{prefix}0x{hwnd:08X}: {class_name} - '{text}' {rect}")
                else:
                    print(f"{prefix}0x{hwnd:08X}: {class_name} {rect}")

                # 자식 컨트롤 재귀
                def enum_child(child_hwnd, param):
                    print_control_tree(child_hwnd, indent + 1)
                    return True

                win32gui.EnumChildWindows(hwnd, enum_child, None)

            except Exception as e:
                pass

        print_control_tree(dlg.handle, 0)

        print("\n=== 2. '기본사항' 텍스트 찾기 ===")

        all_controls = []

        def collect_controls(hwnd, controls_list):
            try:
                class_name = win32gui.GetClassName(hwnd)
                text = win32gui.GetWindowText(hwnd)
                rect = win32gui.GetWindowRect(hwnd)

                controls_list.append({
                    'hwnd': hwnd,
                    'class': class_name,
                    'text': text,
                    'rect': rect
                })

                win32gui.EnumChildWindows(hwnd, lambda h, l: collect_controls(h, l) or True, controls_list)
            except:
                pass
            return True

        win32gui.EnumChildWindows(dlg.handle, lambda h, l: collect_controls(h, l) or True, all_controls)

        # '기본' 또는 '사항' 포함하는 컨트롤 찾기
        basic_controls = [c for c in all_controls if c['text'] and ('기본' in c['text'] or '사항' in c['text'])]

        if basic_controls:
            print(f"'기본사항' 관련 컨트롤 {len(basic_controls)}개 발견:")
            for ctrl in basic_controls:
                print(f"  0x{ctrl['hwnd']:08X}: {ctrl['class']} - '{ctrl['text']}'")
        else:
            print("'기본사항' 텍스트를 찾을 수 없음")

        print("\n=== 3. 탭 컨트롤 상세 분석 ===")

        # 탭 컨트롤만 찾기
        tab_controls = [c for c in all_controls if 'tab' in c['class'].lower()]

        print(f"탭 컨트롤 {len(tab_controls)}개 발견:")
        for tab_ctrl in tab_controls:
            hwnd = tab_ctrl['hwnd']
            print(f"\n  탭: 0x{hwnd:08X} ({tab_ctrl['class']})")

            # TCM_GETITEMCOUNT: 탭 개수
            TCM_GETITEMCOUNT = 0x1304
            count = win32gui.SendMessage(hwnd, TCM_GETITEMCOUNT, 0, 0)
            print(f"    탭 개수: {count}")

            # 각 탭의 텍스트 읽기
            if count > 0:
                TCM_GETITEM = 0x133C  # TCM_GETITEMW
                TCIF_TEXT = 0x0001

                class TCITEM(Structure):
                    _fields_ = [
                        ("mask", c_uint),
                        ("dwState", c_uint),
                        ("dwStateMask", c_uint),
                        ("pszText", c_wchar_p),
                        ("cchTextMax", c_int),
                        ("iImage", c_int),
                        ("lParam", c_void_p)
                    ]

                for i in range(count):
                    buffer = create_unicode_buffer(256)
                    item = TCITEM()
                    item.mask = TCIF_TEXT
                    item.pszText = cast(buffer, c_wchar_p)
                    item.cchTextMax = 256

                    # 프로세스 메모리에 접근해야 하므로 실패할 수 있음
                    try:
                        result = win32gui.SendMessage(hwnd, TCM_GETITEM, i, addressof(item))
                        if result:
                            print(f"      탭 {i}: '{buffer.value}'")
                    except:
                        print(f"      탭 {i}: (읽기 실패)")

        print("\n=== 4. pywinauto로 탭 찾기 ===")

        # pywinauto로 탭 컨트롤 접근
        try:
            # 모든 자식 출력
            print("\npywinauto children():")
            for child in dlg.children():
                class_name = child.class_name()
                try:
                    text = child.window_text()
                except:
                    text = ""

                if text:
                    print(f"  {class_name}: '{text}'")
                elif 'tab' in class_name.lower():
                    print(f"  {class_name}: (탭 컨트롤)")

                    # 탭의 자식들 확인
                    try:
                        for tab_child in child.children():
                            try:
                                tab_text = tab_child.window_text()
                                print(f"    → '{tab_text}'")
                            except:
                                pass
                    except:
                        pass

        except Exception as e:
            print(f"pywinauto 오류: {e}")

        print("\n=== 5. 모든 버튼 찾기 ===")

        buttons = [c for c in all_controls if c['class'] in ['Button', 'Afx:Button', 'TButton'] or 'button' in c['class'].lower()]

        print(f"버튼 컨트롤 {len(buttons)}개 발견:")
        for btn in buttons[:20]:  # 처음 20개만
            if btn['text']:
                print(f"  '{btn['text']}' ({btn['class']})")

        capture_func("attempt72_01_complete.png")

        return {
            "success": True,
            "message": f"""
컨트롤 분석 완료:
- 전체 컨트롤: {len(all_controls)}개
- 탭 컨트롤: {len(tab_controls)}개
- 버튼: {len(buttons)}개
- '기본사항' 관련: {len(basic_controls)}개
"""
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
