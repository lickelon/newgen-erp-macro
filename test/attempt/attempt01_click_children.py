"""
시도 1: 탭 컨트롤의 모든 자식 요소를 click_input()으로 클릭
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
    print("시도 1: 탭 컨트롤 자식 요소 click_input() 테스트")
    print("="*60)

    try:
        # 탭 컨트롤 찾기 (사원등록 프로그램)
        tab_control = dlg.child_window(class_name="Afx:TabWnd:cd0000:8:10003:10", found_index=0)

        if not tab_control.exists():
            return {"success": False, "message": "탭 컨트롤을 찾을 수 없음"}

        rect = tab_control.rectangle()
        print(f"\n탭 컨트롤: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")

        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt01_00_initial.png")

        # 탭 컨트롤의 모든 자식 찾기
        children = tab_control.descendants()
        print(f"\n자식 요소 {len(children)}개 발견")

        # 탭 영역(상단 32px) 내의 요소만 필터링
        tab_children = []
        for child in children:
            try:
                child_rect = child.rectangle()
                # 탭 영역 내에 있는지 확인
                if child_rect.top < rect.top + 32 and child_rect.bottom > rect.top:
                    class_name = child.class_name()
                    tab_children.append((child, class_name, child_rect))
            except:
                pass

        print(f"탭 영역 내 요소: {len(tab_children)}개\n")

        if len(tab_children) == 0:
            return {"success": False, "message": "탭 영역 내 요소를 찾을 수 없음"}

        # 각 요소 클릭 시도
        for i, (child, class_name, child_rect) in enumerate(tab_children, 1):
            print(f"[{i}/{len(tab_children)}] {class_name} 클릭")
            print(f"  위치: L={child_rect.left}, T={child_rect.top}")

            try:
                # click_input() 사용 (마우스 움직이지 않음)
                child.click_input()
                time.sleep(0.3)

                # 스크린샷
                filename = f"attempt01_{i:02d}_{class_name[:20]}.png"
                capture_func(filename)
                print(f"  ✓ 완료")

            except Exception as e:
                print(f"  ✗ 실패: {e}")

            time.sleep(0.3)

        return {"success": True, "message": f"{len(tab_children)}개 요소 클릭 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
