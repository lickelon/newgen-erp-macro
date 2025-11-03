"""
시도 7: 안정적인 탭 컨트롤 찾기 방법 테스트
프로그램 재시작 시 클래스명이 바뀌는지 확인
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
    print("시도 7: 안정적인 탭 컨트롤 찾기")
    print("="*60)

    try:
        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt07_00_initial.png")

        # 방법 1: 기존 클래스명으로 찾기
        print("\n=== 방법 1: 전체 클래스명으로 찾기 ===")
        try:
            tab1 = dlg.child_window(class_name="Afx:TabWnd:cd0000:8:10003:10", found_index=0)
            if tab1.exists():
                print(f"✓ 찾음: {tab1.class_name()}")
                print(f"  HWND: 0x{tab1.handle:08X}")
            else:
                print("✗ 찾을 수 없음")
        except Exception as e:
            print(f"✗ 오류: {e}")

        # 방법 2: 부분 클래스명으로 찾기 (Afx:TabWnd:로 시작)
        print("\n=== 방법 2: 부분 클래스명 매칭 (Afx:TabWnd:) ===")
        descendants = dlg.descendants()
        tab_controls = []

        for ctrl in descendants:
            try:
                class_name = ctrl.class_name()
                if class_name.startswith("Afx:TabWnd:"):
                    tab_controls.append(ctrl)
                    print(f"✓ 찾음: {class_name}")
                    print(f"  HWND: 0x{ctrl.handle:08X}")
                    rect = ctrl.rectangle()
                    print(f"  위치: L={rect.left}, T={rect.top}, R={rect.right}, B={rect.bottom}")
                    print(f"  크기: W={rect.width()}, H={rect.height()}")
            except:
                pass

        # 방법 3: 클래스명 패턴 분석
        print("\n=== 방법 3: Afx로 시작하는 모든 컨트롤 ===")
        afx_controls = []
        for ctrl in descendants:
            try:
                class_name = ctrl.class_name()
                if class_name.startswith("Afx:") and "TabWnd" in class_name:
                    afx_controls.append((ctrl, class_name))
                    print(f"  - {class_name}")
            except:
                pass

        # 방법 4: 부모-자식 관계로 찾기
        print("\n=== 방법 4: 윈도우 계층으로 찾기 ===")
        print(f"부모 윈도우: {dlg.class_name()}")
        print(f"부모 HWND: 0x{dlg.handle:08X}")

        # 직계 자식들 중에서 찾기
        children = dlg.children()
        print(f"\n직계 자식 {len(children)}개:")
        for i, child in enumerate(children):
            try:
                class_name = child.class_name()
                print(f"  [{i}] {class_name}")
                if "TabWnd" in class_name:
                    print(f"      ★ 탭 컨트롤 발견!")
                    print(f"      HWND: 0x{child.handle:08X}")
            except:
                pass

        # 결과 요약
        print("\n=== 결과 요약 ===")
        print(f"부분 매칭으로 찾은 탭 컨트롤: {len(tab_controls)}개")

        if len(tab_controls) > 0:
            print("\n권장 방법:")
            print("1. 'Afx:TabWnd:'로 시작하는 클래스명 검색")
            print("2. found_index=0으로 첫 번째 것 선택")
            print("\n예제 코드:")
            print("```python")
            print("descendants = dlg.descendants()")
            print("for ctrl in descendants:")
            print("    if ctrl.class_name().startswith('Afx:TabWnd:'):")
            print("        tab_control = ctrl")
            print("        break")
            print("```")

            return {"success": True, "message": f"{len(tab_controls)}개 탭 컨트롤 발견"}
        else:
            return {"success": False, "message": "탭 컨트롤을 찾을 수 없음"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
