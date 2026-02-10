"""
시도 90: 분납적용 창 찾기 및 구조 분석
"""
import time

def run(dlg, capture_func):
    """
    Args:
        dlg: pywinauto 윈도우 객체 (분납적용)
        capture_func: 스크린샷 함수

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 90: 분납적용 창 찾기 및 구조 분석")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt90_00_initial.png")

        # 기본 정보
        print(f"\n[기본 정보]")
        print(f"  제목: {dlg.window_text()}")
        print(f"  클래스: {dlg.class_name()}")
        print(f"  HWND: 0x{dlg.handle:08X}")

        rect = dlg.rectangle()
        print(f"  위치: ({rect.left}, {rect.top})")
        print(f"  크기: {rect.width()} x {rect.height()}")

        # 자식 컨트롤
        children = dlg.children()
        print(f"\n[자식 컨트롤]: {len(children)}개")

        all_controls = dlg.descendants()
        print(f"[전체 컨트롤]: {len(all_controls)}개")

        # 타입별 분류
        control_types = {}
        for ctrl in all_controls:
            try:
                cn = ctrl.class_name()
                if cn not in control_types:
                    control_types[cn] = []
                control_types[cn].append(ctrl)
            except:
                pass

        print(f"\n[컨트롤 타입별]")
        for cn, ctrls in sorted(control_types.items()):
            print(f"  {cn}: {len(ctrls)}개")

        # 중요 컨트롤 상세
        for ctrl_type in ['fpUSpread80', 'Button', 'Edit', 'ComboBox', 'Static']:
            if ctrl_type in control_types:
                print(f"\n[{ctrl_type}]")
                for i, ctrl in enumerate(control_types[ctrl_type][:5]):
                    try:
                        text = ctrl.window_text()
                        print(f"  #{i+1} 0x{ctrl.handle:08X} '{text}'")
                    except:
                        pass

        return {"success": True, "message": "구조 분석 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
