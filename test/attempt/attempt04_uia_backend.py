"""
시도 4: UIA (UI Automation) 백엔드로 탭 선택
Win32 대신 UIA 백엔드를 사용하여 탭 접근
"""
import time

def run(dlg_win32, capture_func):
    """
    Args:
        dlg_win32: pywinauto 윈도우 객체 (win32)
        capture_func: 스크린샷 함수 (filename) -> None

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 4: UIA 백엔드로 탭 선택")
    print("="*60)

    try:
        from pywinauto import application

        # 초기 상태
        print("\n📸 초기 상태 캡처")
        capture_func("attempt04_00_initial.png")

        # UIA 백엔드로 재연결
        print("\nUIA 백엔드로 연결 중...")
        app_uia = application.Application(backend="uia")
        app_uia.connect(title="사원등록")
        dlg_uia = app_uia.window(title="사원등록")

        print("✓ UIA 백엔드 연결 성공")

        # UIA로 탭 컨트롤 찾기
        print("\n탭 컨트롤 검색 중...")

        # 모든 컨트롤 출력
        descendants = dlg_uia.descendants()
        print(f"총 {len(descendants)}개 컨트롤 발견")

        # Tab 관련 컨트롤 찾기
        tab_controls = []
        for ctrl in descendants:
            try:
                ctrl_type = ctrl.element_info.control_type
                name = ctrl.element_info.name

                if "Tab" in ctrl_type or "tab" in ctrl_type.lower():
                    print(f"✓ 탭 컨트롤 발견: {ctrl_type}, name={name}")
                    tab_controls.append(ctrl)

                # 탭 아이템도 찾기
                if "TabItem" in ctrl_type:
                    print(f"  - 탭 아이템: {name}")

            except:
                pass

        if len(tab_controls) == 0:
            return {"success": False, "message": "UIA로 탭 컨트롤을 찾을 수 없음"}

        # 첫 번째 탭 컨트롤 사용
        tab_control = tab_controls[0]
        print(f"\n탭 컨트롤 선택: {tab_control}")

        # 탭 아이템들 찾기
        print("\n탭 아이템 검색 중...")
        tab_items = []
        for child in tab_control.descendants():
            try:
                ctrl_type = child.element_info.control_type
                if "TabItem" in ctrl_type:
                    name = child.element_info.name
                    tab_items.append((child, name))
                    print(f"  - 탭 아이템 {len(tab_items)}: {name}")
            except:
                pass

        if len(tab_items) == 0:
            return {"success": False, "message": "탭 아이템을 찾을 수 없음"}

        # 각 탭 아이템 선택 시도
        print(f"\n{len(tab_items)}개 탭 아이템 선택 시도...")
        for i, (tab_item, name) in enumerate(tab_items, 1):
            print(f"\n[{i}/{len(tab_items)}] '{name}' 탭 선택 중...")

            try:
                # UIA의 select 패턴 사용
                tab_item.select()
                time.sleep(0.5)

                capture_func(f"attempt04_01_select_{i}_{name}.png")
                print(f"  ✓ 완료")

            except Exception as e:
                print(f"  ✗ 실패: {e}")

        return {"success": True, "message": f"{len(tab_items)}개 탭 아이템 선택 완료"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
