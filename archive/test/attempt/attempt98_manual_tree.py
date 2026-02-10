"""
시도 98: 수동으로 컨트롤 트리 출력
"""
from pywinauto import Application

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 98: 수동 컨트롤 트리 출력")
    print("="*60)

    try:
        # 급여자료입력 연결
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")

        print(f"\n급여자료입력 HWND: 0x{main_window.handle:08X}")

        # 모든 하위 컨트롤 가져오기
        print("\n모든 하위 컨트롤 수집 중...")
        all_controls = main_window.descendants()
        print(f"총 {len(all_controls)}개 컨트롤")

        # 파일로 저장
        output_lines = []
        output_lines.append(f"급여자료입력 컨트롤 트리")
        output_lines.append(f"HWND: 0x{main_window.handle:08X}")
        output_lines.append("=" * 80)
        output_lines.append("")

        found_installment = False

        for i, ctrl in enumerate(all_controls):
            try:
                class_name = ctrl.class_name()
                title = ctrl.window_text()
                hwnd = ctrl.handle

                # 중요한 컨트롤만 출력
                if class_name in ["#32770", "Button"] or title:
                    line = f"[{i+1}] 0x{hwnd:08X} | {class_name:30s}"
                    if title:
                        line += f" | '{title}'"

                    output_lines.append(line)
                    print(line)

                    if "분납" in str(title):
                        output_lines.append(f"      >>> ⭐⭐⭐ 분납 관련 발견!")
                        print(f"      >>> ⭐⭐⭐ 분납 관련 발견!")
                        found_installment = True
            except:
                pass

        # 파일로 저장
        with open("test/control_tree.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))

        print(f"\n✓ 저장: test/control_tree.txt ({len(output_lines)} lines)")

        if found_installment:
            return {"success": True, "message": "분납 관련 컨트롤 발견"}
        else:
            return {"success": False, "message": "분납 관련 컨트롤을 찾지 못했습니다"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
