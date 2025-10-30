"""
Attempt 25: 부양가족명세 버튼/컨트롤 찾아서 클릭

좌표 없이 컨트롤 이름/텍스트로 찾아서 클릭
"""
import time


def find_and_click_dependents(dlg, capture_func):
    """부양가족명세 관련 컨트롤 찾아서 클릭"""
    print("\n" + "=" * 60)
    print("Attempt 25: 부양가족명세 클릭")
    print("=" * 60)

    try:
        capture_func("attempt25_00_initial.png")

        # 기본사항 탭
        print("\n[1/4] 기본사항 탭 선택...")
        from tab_automation import TabAutomation
        tab_auto = TabAutomation()
        tab_auto.connect()
        tab_auto.select_tab("기본사항")
        time.sleep(0.5)
        capture_func("attempt25_01_basic_tab.png")

        # 모든 컨트롤 탐색
        print("\n[2/4] 부양가족 관련 컨트롤 찾기...")

        dependent_controls = []
        for ctrl in dlg.descendants():
            try:
                # 컨트롤의 텍스트 확인
                text = ctrl.window_text()
                class_name = ctrl.class_name()

                # "부양가족" 포함된 것들 찾기
                if "부양" in text or "가족" in text or "명세" in text:
                    dependent_controls.append({
                        "ctrl": ctrl,
                        "text": text,
                        "class": class_name,
                        "visible": ctrl.is_visible(),
                        "enabled": ctrl.is_enabled()
                    })
            except:
                pass

        print(f"  부양가족 관련 컨트롤 {len(dependent_controls)}개 발견:")
        for idx, info in enumerate(dependent_controls[:10]):  # 최대 10개만
            visible = "✓" if info["visible"] else "✗"
            enabled = "✓" if info["enabled"] else "✗"
            print(f"    [{idx}] {info['class']}")
            print(f"        텍스트: \"{info['text']}\"")
            print(f"        visible: {visible}, enabled: {enabled}")

        # 클릭 가능한 버튼 찾기
        print("\n[3/4] 클릭 가능한 버튼 찾기...")

        clickable_buttons = []
        for info in dependent_controls:
            if info["visible"] and info["enabled"]:
                class_name = info["class"]
                # 버튼 계열 클래스
                if "Button" in class_name or "Btn" in class_name or class_name == "ToolbarWindow32":
                    clickable_buttons.append(info)

        print(f"  클릭 가능한 버튼: {len(clickable_buttons)}개")
        for idx, info in enumerate(clickable_buttons):
            print(f"    [{idx}] {info['class']}: \"{info['text']}\"")

        # 첫 번째 버튼 클릭 시도
        if len(clickable_buttons) > 0:
            print("\n[4/4] 첫 번째 버튼 클릭 시도...")
            target = clickable_buttons[0]
            print(f"  타겟: {target['class']} - \"{target['text']}\"")

            try:
                target["ctrl"].click_input()
                time.sleep(0.5)
                capture_func("attempt25_02_after_click.png")
                print("  ✓ 클릭 완료")

                return {
                    "success": True,
                    "message": f"버튼 클릭 성공: {target['text']}",
                    "button_text": target['text']
                }
            except Exception as e:
                print(f"  ✗ 클릭 실패: {e}")
                return {"success": False, "message": f"클릭 실패: {e}"}

        # 버튼이 없으면 모든 부양가족 컨트롤 정보 반환
        capture_func("attempt25_03_no_button.png")
        return {
            "success": False,
            "message": "클릭 가능한 버튼 없음",
            "total_controls": len(dependent_controls),
            "controls": [f"{c['class']}: {c['text']}" for c in dependent_controls[:5]]
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

    result = find_and_click_dependents(dlg, capture_func)
    print(f"\n{'='*60}")
    print(f"최종 결과: {result}")
    print(f"{'='*60}")
