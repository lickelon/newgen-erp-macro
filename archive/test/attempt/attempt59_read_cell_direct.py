"""
시도 59: 복사 없이 현재 셀 값 직접 읽기

클립보드를 사용하지 않고 fpUSpread80의 현재 선택된 셀 값을
직접 읽어오는 방법 탐색
"""
import time
import win32api
import win32con
import win32gui
from ctypes import windll, create_unicode_buffer, c_int, byref


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 59: 복사 없이 현재 셀 값 직접 읽기")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt59_00_initial.png")

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

        # 참조 값 (복사로 확인)
        import pyperclip
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (복사로 확인): '{reference_value}'")

        print("\n=== 방법 1: 자식 윈도우에서 텍스트 읽기 ===")
        try:
            # 자식 윈도우 열거
            def enum_child_callback(hwnd_child, results):
                try:
                    class_name = win32gui.GetClassName(hwnd_child)

                    # 텍스트 길이 확인
                    length = win32gui.SendMessage(hwnd_child, win32con.WM_GETTEXTLENGTH, 0, 0)

                    if length > 0:
                        # 텍스트 읽기
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(hwnd_child, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value

                        results.append({
                            'hwnd': hwnd_child,
                            'class': class_name,
                            'text': text
                        })
                        print(f"  자식 0x{hwnd_child:08X} ({class_name}): '{text}'")
                except:
                    pass
                return True

            children_with_text = []
            win32gui.EnumChildWindows(hwnd, enum_child_callback, children_with_text)

            if children_with_text:
                print(f"✓ 텍스트가 있는 자식 윈도우 {len(children_with_text)}개 발견")
                # 참조 값과 일치하는지 확인
                for child in children_with_text:
                    if reference_value in child['text'] or child['text'] in reference_value:
                        print(f"  ✓ 일치하는 자식 발견! HWND=0x{child['hwnd']:08X}, 값='{child['text']}'")
            else:
                print("✗ 텍스트가 있는 자식 윈도우 없음")

        except Exception as e:
            import traceback
            print(f"✗ 방법 1 실패: {e}")
            traceback.print_exc()

        print("\n=== 방법 2: 윈도우 프로퍼티 확인 ===")
        try:
            # 알려진 프로퍼티 이름들 시도
            prop_names = [
                "CellText",
                "CellValue",
                "CurrentCell",
                "SelectedCell",
                "Text",
                "Value",
                "Data",
            ]

            for prop_name in prop_names:
                try:
                    # GetProp 시도
                    atom = win32api.GlobalFindAtom(prop_name)
                    if atom:
                        prop_value = win32gui.GetProp(hwnd, atom)
                        if prop_value:
                            print(f"  {prop_name}: 0x{prop_value:08X}")
                except:
                    pass

        except Exception as e:
            print(f"✗ 방법 2 실패: {e}")

        print("\n=== 방법 3: GetWindowLongPtr로 인스턴스 데이터 접근 ===")
        try:
            # GWLP_USERDATA 확인
            user_data = win32gui.GetWindowLong(hwnd, win32con.GWL_USERDATA)
            print(f"GWL_USERDATA: 0x{user_data:08X}")

            if user_data != 0:
                print("  → 0이 아닌 값 발견, 하지만 직접 역참조는 불가능")

            # 다른 인덱스들 확인
            indices = [
                (win32con.GWL_WNDPROC, "GWL_WNDPROC"),
                (win32con.GWL_HINSTANCE, "GWL_HINSTANCE"),
                (win32con.GWL_ID, "GWL_ID"),
                (win32con.GWL_STYLE, "GWL_STYLE"),
                (win32con.GWL_EXSTYLE, "GWL_EXSTYLE"),
            ]

            for idx, name in indices:
                try:
                    value = win32gui.GetWindowLong(hwnd, idx)
                    print(f"  {name}: 0x{value:08X}")
                except:
                    pass

        except Exception as e:
            print(f"✗ 방법 3 실패: {e}")

        print("\n=== 방법 4: 특수 메시지 시도 ===")
        try:
            # Spread 컨트롤의 일반적인 메시지들
            messages = [
                # 텍스트 관련
                (win32con.WM_GETTEXT, "WM_GETTEXT"),
                (0x400, "WM_USER"),
                (0x400 + 1, "WM_USER+1"),
                (0x400 + 100, "WM_USER+100"),
                (0x400 + 200, "WM_USER+200"),
                # 에디트 컨트롤 메시지
                (0xB0, "EM_GETSEL"),
                (0xC0, "EM_GETLINE"),
                (0xCE, "EM_GETLINECOUNT"),
            ]

            for msg, name in messages:
                try:
                    if msg == win32con.WM_GETTEXT:
                        # 이미 시도했지만 다시 한번
                        buffer = create_unicode_buffer(1024)
                        result = win32gui.SendMessage(hwnd, msg, 1024, buffer)
                        if result > 0:
                            text = buffer.value
                            print(f"  {name}: '{text}' (len={result})")
                    elif msg == 0xCE:  # EM_GETLINECOUNT
                        result = win32gui.SendMessage(hwnd, msg, 0, 0)
                        if result > 0:
                            print(f"  {name}: {result} lines")
                    else:
                        result = win32gui.SendMessage(hwnd, msg, 0, 0)
                        if result != 0:
                            print(f"  {name}: {result} (0x{result:08X})")
                except:
                    pass

        except Exception as e:
            print(f"✗ 방법 4 실패: {e}")

        print("\n=== 방법 5: 포커스된 자식 윈도우 확인 ===")
        try:
            # 현재 포커스된 윈도우 확인
            focused = win32gui.GetFocus()
            print(f"현재 포커스: 0x{focused:08X}")

            if focused != hwnd:
                # 포커스된 윈도우가 다르면 거기서 텍스트 읽기
                focused_class = win32gui.GetClassName(focused)
                print(f"포커스 클래스: {focused_class}")

                length = win32gui.SendMessage(focused, win32con.WM_GETTEXTLENGTH, 0, 0)
                if length > 0:
                    buffer = create_unicode_buffer(length + 1)
                    win32gui.SendMessage(focused, win32con.WM_GETTEXT, length + 1, buffer)
                    text = buffer.value
                    print(f"✓ 포커스된 윈도우 텍스트: '{text}'")

                    if text == reference_value:
                        print(f"  ✓✓ 참조 값과 일치!")
                else:
                    print("✗ 포커스된 윈도우에 텍스트 없음")
            else:
                print("  (fpUSpread80 자체가 포커스됨)")

        except Exception as e:
            print(f"✗ 방법 5 실패: {e}")

        print("\n=== 방법 6: 모든 관련 윈도우 트리 탐색 ===")
        try:
            # 부모-자식 관계 전체 출력
            def print_window_tree(hwnd_root, indent=0):
                try:
                    class_name = win32gui.GetClassName(hwnd_root)
                    length = win32gui.SendMessage(hwnd_root, win32con.WM_GETTEXTLENGTH, 0, 0)

                    info = f"{'  ' * indent}0x{hwnd_root:08X} ({class_name})"

                    if length > 0:
                        buffer = create_unicode_buffer(length + 1)
                        win32gui.SendMessage(hwnd_root, win32con.WM_GETTEXT, length + 1, buffer)
                        text = buffer.value
                        if text:
                            info += f" = '{text[:50]}'"

                    print(info)

                    # 자식 윈도우 재귀
                    def enum_callback(hwnd_child, param):
                        print_window_tree(hwnd_child, indent + 1)
                        return True

                    win32gui.EnumChildWindows(hwnd_root, enum_callback, None)
                except:
                    pass

            print("윈도우 트리:")
            print_window_tree(hwnd)

        except Exception as e:
            print(f"✗ 방법 6 실패: {e}")

        capture_func("attempt59_01_complete.png")

        return {
            "success": False,
            "message": "셀 값 직접 읽기 실패 - 추가 조사 필요"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
