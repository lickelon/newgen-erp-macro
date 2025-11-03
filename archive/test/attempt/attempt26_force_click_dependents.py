"""
Attempt 26: 부양가족명세 강제 클릭 (visible=False여도)

hidden 컨트롤도 클릭 시도
"""
import time


def force_click_dependents(dlg, capture_func):
    """부양가족명세 컨트롤 강제 클릭"""
    print("\n" + "=" * 60)
    print("Attempt 26: 부양가족명세 강제 클릭")
    print("=" * 60)

    try:
        capture_func("attempt26_00_initial.png")

        # 기본사항 탭
        print("\n[1/3] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt26_01_basic_tab.png")

        # 부양가족명세 컨트롤 찾기
        print("\n[2/3] 부양가족명세 컨트롤 찾기...")

        target_ctrl = None
        for ctrl in dlg.descendants():
            try:
                text = ctrl.window_text()
                if "부양가족명세" in text:
                    print(f"  발견: {ctrl.class_name()}")
                    print(f"    텍스트: \"{text}\"")
                    print(f"    visible: {ctrl.is_visible()}")
                    print(f"    enabled: {ctrl.is_enabled()}")
                    print(f"    HWND: 0x{ctrl.handle:08X}")
                    target_ctrl = ctrl
                    break
            except:
                pass

        if not target_ctrl:
            return {"success": False, "message": "부양가족명세 컨트롤 없음"}

        # 강제 클릭 시도
        print("\n[3/3] 강제 클릭 시도...")

        methods = [
            ("click_input()", lambda: target_ctrl.click_input()),
            ("click()", lambda: target_ctrl.click()),
            ("SendMessage(BM_CLICK)", lambda: target_ctrl.send_message(0x00F5, 0, 0)),  # BM_CLICK
        ]

        for method_name, method_func in methods:
            try:
                print(f"  시도: {method_name}")
                method_func()
                time.sleep(0.5)
                capture_func(f"attempt26_02_after_{method_name.replace('()', '')}.png")
                print(f"    ✓ {method_name} 완료")

                return {
                    "success": True,
                    "message": f"{method_name} 성공",
                    "method": method_name
                }
            except Exception as e:
                print(f"    ✗ {method_name} 실패: {e}")

        capture_func("attempt26_03_all_failed.png")
        return {"success": False, "message": "모든 클릭 방법 실패"}

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

    result = force_click_dependents(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
