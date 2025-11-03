"""
Attempt 37: 탭 선택만 테스트

기본사항 → 부양가족정보 탭 전환이 제대로 되는지만 확인
"""
import time


def test_tab_selection(dlg, capture_func):
    """탭 선택만 테스트"""
    print("\n" + "=" * 60)
    print("Attempt 37: 탭 선택 테스트")
    print("=" * 60)

    try:
        from tab_automation import TabAutomation

        # 현재 상태
        print("\n[1/4] 현재 상태 확인...")
        capture_func("attempt37_00_initial.png")
        time.sleep(0.5)

        tab_auto = TabAutomation()
        tab_auto.connect()

        # 기본사항 탭
        print("\n[2/4] 기본사항 탭 선택...")
        try:
            result = tab_auto.select_tab("기본사항")
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt37_01_basic_tab.png")
        except Exception as e:
            print(f"  ✗ 기본사항 탭 선택 실패: {e}")
            capture_func("attempt37_01_basic_tab_failed.png")

        # 부양가족정보 탭
        print("\n[3/4] 부양가족정보 탭 선택...")
        try:
            result = tab_auto.select_tab("부양가족정보")
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt37_02_dependents_tab.png")
        except Exception as e:
            print(f"  ✗ 부양가족정보 탭 선택 실패: {e}")
            capture_func("attempt37_02_dependents_tab_failed.png")

        # 다시 기본사항
        print("\n[4/4] 다시 기본사항 탭 선택...")
        try:
            result = tab_auto.select_tab("기본사항")
            print(f"  결과: {result}")
            time.sleep(1.0)
            capture_func("attempt37_03_basic_tab_again.png")
        except Exception as e:
            print(f"  ✗ 기본사항 탭 재선택 실패: {e}")
            capture_func("attempt37_03_basic_tab_again_failed.png")

        print("\n" + "=" * 60)
        print("탭 선택 테스트 완료")
        print("=" * 60)

        return {
            "success": True,
            "message": "탭 선택 테스트 완료"
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

    result = test_tab_selection(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
