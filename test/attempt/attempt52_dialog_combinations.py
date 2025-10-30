"""
Attempt 52: 다이얼로그 제어 다양한 조합 시도

ShowWindow, SetWindowPos, EnableWindow 등 다양한 조합
"""
import time
import win32gui
import win32con


def test_dialog_combinations(dlg, capture_func):
    """다양한 조합 시도"""
    print("\n" + "=" * 60)
    print("Attempt 52: 다이얼로그 다양한 조합")
    print("=" * 60)

    try:
        # 현재 상태
        print("\n[초기] 현재 상태...")
        capture_func("attempt52_00_initial.png")
        time.sleep(0.5)

        # 다이얼로그 찾기
        print("\n다이얼로그 검색...")
        basic_dialog = None
        dependents_dialog = None

        for ctrl in dlg.descendants():
            try:
                if ctrl.class_name() == "#32770":
                    text = ctrl.window_text().strip()
                    if "기본사항" in text:
                        basic_dialog = ctrl
                        print(f"  ✓ 기본사항: 0x{ctrl.handle:08X}")
                    elif "부양가족명세" in text:
                        dependents_dialog = ctrl
                        print(f"  ✓ 부양가족명세: 0x{ctrl.handle:08X}")
            except:
                pass

        if not basic_dialog or not dependents_dialog:
            return {"success": False, "message": "다이얼로그 없음"}

        # 조합 1: 부양가족 숨기고 기본사항 보이기
        print("\n[1] ShowWindow(SW_HIDE) + ShowWindow(SW_SHOW)")
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.3)
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_SHOW)
        time.sleep(1.0)
        capture_func("attempt52_01_hide_show.png")

        # 조합 2: BringWindowToTop + SetForegroundWindow
        print("\n[2] BringWindowToTop + SetForegroundWindow")
        win32gui.BringWindowToTop(basic_dialog.handle)
        time.sleep(0.2)
        try:
            win32gui.SetForegroundWindow(basic_dialog.handle)
        except:
            print("  (SetForegroundWindow 실패 - 무시)")
        time.sleep(1.0)
        capture_func("attempt52_02_bring_setfg.png")

        # 조합 3: 부양가족으로 전환 (숨기고 보이기)
        print("\n[3] 부양가족명세로 전환 (SW_HIDE + SW_SHOW)")
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.3)
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_SHOW)
        time.sleep(1.0)
        capture_func("attempt52_03_dependents_hide_show.png")

        # 조합 4: SetWindowPos로 Z-order 변경
        print("\n[4] SetWindowPos (HWND_TOP)")
        win32gui.SetWindowPos(
            dependents_dialog.handle,
            win32con.HWND_TOP,
            0, 0, 0, 0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        )
        time.sleep(1.0)
        capture_func("attempt52_04_setwindowpos.png")

        # 조합 5: EnableWindow + ShowWindow
        print("\n[5] 기본사항 EnableWindow + ShowWindow")
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.2)
        win32gui.EnableWindow(basic_dialog.handle, True)
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_SHOW)
        time.sleep(1.0)
        capture_func("attempt52_05_enable_show.png")

        # 조합 6: 다시 부양가족 (EnableWindow + SetWindowPos)
        print("\n[6] 부양가족명세 EnableWindow + SetWindowPos")
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.2)
        win32gui.EnableWindow(dependents_dialog.handle, True)
        win32gui.SetWindowPos(
            dependents_dialog.handle,
            win32con.HWND_TOP,
            0, 0, 0, 0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        )
        time.sleep(1.0)
        capture_func("attempt52_06_dependents_enable_setpos.png")

        # 조합 7: pywinauto의 set_focus + ShowWindow
        print("\n[7] pywinauto set_focus + ShowWindow")
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.2)
        basic_dialog.set_focus()
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_SHOW)
        time.sleep(1.0)
        capture_func("attempt52_07_setfocus_show.png")

        # 조합 8: 다시 부양가족 (set_focus + ShowWindow)
        print("\n[8] 부양가족명세 set_focus + ShowWindow")
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.2)
        dependents_dialog.set_focus()
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_SHOW)
        time.sleep(1.0)
        capture_func("attempt52_08_dependents_setfocus_show.png")

        # 조합 9: InvalidateRect + UpdateWindow
        print("\n[9] 기본사항 InvalidateRect + UpdateWindow")
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.2)
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_SHOW)
        win32gui.InvalidateRect(basic_dialog.handle, None, True)
        win32gui.UpdateWindow(basic_dialog.handle)
        time.sleep(1.0)
        capture_func("attempt52_09_invalidate_update.png")

        # 조합 10: 다시 부양가족 최종
        print("\n[10] 부양가족명세 최종")
        win32gui.ShowWindow(basic_dialog.handle, win32con.SW_HIDE)
        time.sleep(0.2)
        win32gui.ShowWindow(dependents_dialog.handle, win32con.SW_SHOW)
        win32gui.SetWindowPos(
            dependents_dialog.handle,
            win32con.HWND_TOP,
            0, 0, 0, 0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        )
        time.sleep(1.0)
        capture_func("attempt52_10_final.png")

        print("\n" + "=" * 60)
        print("완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "다양한 조합 테스트 완료"
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

    result = test_dialog_combinations(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
