"""
시도 2: 탭 컨트롤 검색 - 모든 컨트롤 스캔하여 Tab 관련 클래스 찾기
"""
import time

def run(dlg, capture_func):
    """
    Args:
        dlg: pywinauto 윈도우 객체
        capture_func: 스크린샷 함수 (filename) -> None

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 2: 탭 컨트롤 검색")
    print("="*60)

    try:
        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt02_00_initial.png")

        # 모든 자식 컨트롤 검색
        print("\n모든 컨트롤 검색 중...")
        descendants = dlg.descendants()

        print(f"\n총 {len(descendants)}개 컨트롤 발견\n")

        # Tab 관련 컨트롤 필터링
        tab_controls = []
        for ctrl in descendants:
            try:
                class_name = ctrl.class_name()
                if "Tab" in class_name or "tab" in class_name.lower():
                    rect = ctrl.rectangle()
                    tab_controls.append((ctrl, class_name, rect))
                    print(f"✓ 탭 컨트롤 발견: {class_name}")
                    print(f"  위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")
                    print(f"  크기: W={rect.width()}, H={rect.height()}")
                    print()
            except:
                pass

        if len(tab_controls) == 0:
            print("⚠️  Tab 관련 컨트롤을 찾을 수 없습니다.")
            print("\n다른 패턴으로 검색 중...")

            # Afx로 시작하는 커스텀 컨트롤 검색
            custom_controls = []
            for ctrl in descendants:
                try:
                    class_name = ctrl.class_name()
                    if class_name.startswith("Afx:") and "Wnd" in class_name:
                        rect = ctrl.rectangle()
                        # 너비가 넓고 높이가 낮은 컨트롤 (탭처럼 생긴 것)
                        if rect.height() < 50 and rect.width() > 200:
                            custom_controls.append((ctrl, class_name, rect))
                            print(f"? 의심 컨트롤: {class_name}")
                            print(f"  위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")
                            print(f"  크기: W={rect.width()}, H={rect.height()}")
                            print()
                except:
                    pass

            if len(custom_controls) > 0:
                print(f"\n{len(custom_controls)}개의 의심 컨트롤 발견")
                return {"success": True, "message": f"{len(custom_controls)}개 의심 컨트롤 발견"}
            else:
                return {"success": False, "message": "탭 컨트롤을 찾을 수 없음"}

        return {"success": True, "message": f"{len(tab_controls)}개 탭 컨트롤 발견"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
