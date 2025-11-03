"""
Attempt 51: 부양가족명세 다이얼로그를 직접 선택

result.txt에서 발견한 #32770 다이얼로그를 pywinauto로 직접 접근
"""
import time
import win32gui


def test_select_by_dialog(dlg, capture_func):
    """다이얼로그 직접 선택"""
    print("\n" + "=" * 60)
    print("Attempt 51: 다이얼로그 직접 선택")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[1/5] 현재 상태...")
        capture_func("attempt51_00_initial.png")
        time.sleep(0.5)

        # 부양가족명세 다이얼로그 찾기
        print("\n[2/5] '부양가족명세' 다이얼로그 찾기...")
        dependents_dialog = None

        # descendants()로 직접 검색
        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "#32770":
                    text = ctrl.window_text().strip()
                    if "부양가족명세" in text:
                        dependents_dialog = ctrl
                        print(f"  ✓ 찾음: '{text}'")
                        print(f"     HWND: 0x{ctrl.handle:08X}")
                        print(f"     Visible: {ctrl.is_visible()}")
                        print(f"     Enabled: {ctrl.is_enabled()}")
                        break
            except:
                pass

        if not dependents_dialog:
            print("  ✗ 부양가족명세 다이얼로그를 찾지 못함")
            return {"success": False, "message": "다이얼로그 없음"}

        # 다이얼로그 활성화 시도 - set_focus
        print("\n[3/5] set_focus() 시도...")
        try:
            dependents_dialog.set_focus()
            print("  ✓ set_focus() 성공")
            time.sleep(1.0)
            capture_func("attempt51_01_after_set_focus.png")
        except Exception as e:
            print(f"  ✗ set_focus() 실패: {e}")

        # ShowWindow로 표시 시도
        print("\n[4/5] ShowWindow() 시도...")
        try:
            import win32con
            hwnd = dependents_dialog.handle
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            print("  ✓ ShowWindow(SW_SHOW) 호출")
            time.sleep(1.0)
            capture_func("attempt51_02_after_show_window.png")
        except Exception as e:
            print(f"  ✗ ShowWindow() 실패: {e}")

        # BringWindowToTop 시도
        print("\n[5/5] BringWindowToTop() 시도...")
        try:
            win32gui.BringWindowToTop(hwnd)
            print("  ✓ BringWindowToTop() 호출")
            time.sleep(1.0)
            capture_func("attempt51_03_after_bring_to_top.png")
        except Exception as e:
            print(f"  ✗ BringWindowToTop() 실패: {e}")

        # 기본사항 다이얼로그로 전환 시도
        print("\n[추가] '기본사항' 다이얼로그로 전환...")
        try:
            basic_dialog = None
            for ctrl in dlg.descendants():
                try:
                    if ctrl.class_name() == "#32770" and "기본사항" in ctrl.window_text():
                        basic_dialog = ctrl
                        break
                except:
                    pass

            if not basic_dialog:
                print("  ✗ 기본사항 다이얼로그 없음")
            else:
                print(f"  ✓ 기본사항 찾음: 0x{basic_dialog.handle:08X}")

            win32gui.ShowWindow(basic_dialog.handle, win32con.SW_SHOW)
            win32gui.BringWindowToTop(basic_dialog.handle)
            time.sleep(1.0)
            capture_func("attempt51_04_basic_dialog.png")
        except Exception as e:
            print(f"  ✗ 기본사항 전환 실패: {e}")

        # 다시 부양가족명세로
        print("\n[추가] 다시 '부양가족명세'로...")
        try:
            win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_SHOW)
            win32gui.BringWindowToTop(dependents_dialog.handle)
            time.sleep(1.0)
            capture_func("attempt51_05_dependents_again.png")
        except Exception as e:
            print(f"  ✗ 부양가족명세 재전환 실패: {e}")

        capture_func("attempt51_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "다이얼로그 직접 선택 완료"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }


if __name__ == "__main__":
    from pywinauto import application
    import sys
    import os

    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    from test.capture import capture_window

    app = application.Application(backend="win32")
    app.connect(title="사원등록")
    dlg = app.window(title="사원등록")

    def capture_func(filename):
        capture_window(dlg.handle, filename)

    result = test_select_by_dialog(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
