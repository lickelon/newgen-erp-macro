"""
Attempt 27: 부양가족명세 화면으로 이동

기본사항 탭 선택 없이, 부양가족명세 컨트롤을 클릭하여 이동
"""
import time


def navigate_to_dependents(dlg, capture_func):
    """부양가족명세 화면으로 이동"""
    print("\n" + "=" * 60)
    print("Attempt 27: 부양가족명세로 이동")
    print("=" * 60)

    try:
        capture_func("attempt27_00_initial.png")

        # 부양가족명세 컨트롤 찾기 (탭 전환 없이!)
        print("\n[1/2] 부양가족명세 컨트롤 찾기...")

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

        # 클릭하여 이동
        print("\n[2/2] 부양가족명세로 이동...")

        try:
            target_ctrl.click_input()
            time.sleep(0.5)
            capture_func("attempt27_01_after_click.png")
            print("  ✓ 클릭 완료")

            # 화면 전환 확인을 위해 추가 대기
            time.sleep(1)
            capture_func("attempt27_02_final.png")

            return {
                "success": True,
                "message": "부양가족명세로 이동 시도 완료",
            }
        except Exception as e:
            print(f"  ✗ 클릭 실패: {e}")
            return {"success": False, "message": f"클릭 실패: {e}"}

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

    result = navigate_to_dependents(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
