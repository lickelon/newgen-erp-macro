"""
기본사항 탭 컨트롤 구조 분석

사번, 성명, 주민번호, 나이 입력 필드를 찾습니다.
"""
import sys
from pywinauto import application

# UTF-8 출력
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')


def analyze_basic_tab():
    print("=" * 80)
    print("기본사항 탭 컨트롤 구조 분석")
    print("=" * 80)

    # 연결
    print("\n[1/3] 사원등록 윈도우 연결 중...")
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        print(f"✓ 연결 성공: HWND=0x{dlg.handle:08X}")
    except Exception as e:
        print(f"✗ 연결 실패: {e}")
        return

    # 기본사항 탭으로 이동
    print("\n[2/3] 기본사항 탭 선택...")
    from tab_automation import TabAutomation
    tab_auto = TabAutomation()
    tab_auto.connect()
    tab_auto.select_tab("기본사항")
    print("✓ 기본사항 탭 선택됨")

    # 컨트롤 분석
    print("\n[3/3] 컨트롤 구조 분석...")
    print("-" * 80)

    descendants = dlg.descendants()
    print(f"총 {len(descendants)}개 컨트롤 발견\n")

    # Edit 컨트롤 필터링 (입력 필드)
    edit_controls = []
    for i, ctrl in enumerate(descendants):
        try:
            class_name = ctrl.class_name()
            if "Edit" in class_name or "edit" in class_name:
                edit_controls.append((i, ctrl))
        except:
            pass

    print(f"📝 Edit 컨트롤 {len(edit_controls)}개 발견:\n")
    for idx, (orig_idx, ctrl) in enumerate(edit_controls):
        try:
            hwnd = ctrl.handle
            class_name = ctrl.class_name()
            rect = ctrl.rectangle()

            # 텍스트 가져오기 시도
            try:
                text = ctrl.window_text()
                text_str = f' "{text}"' if text else " (empty)"
            except:
                text_str = " (no text)"

            # 가시성 확인
            visible = ctrl.is_visible()
            enabled = ctrl.is_enabled()

            print(f"  [{idx}] HWND=0x{hwnd:08X} {class_name}")
            print(f"      위치: L={rect.left} T={rect.top} R={rect.right} B={rect.bottom}")
            print(f"      크기: W={rect.width()} H={rect.height()}")
            print(f"      텍스트:{text_str}")
            print(f"      상태: visible={visible}, enabled={enabled}")
            print()

        except Exception as e:
            print(f"  [{idx}] 오류: {e}")
            print()

    # 기본사항 다이얼로그 찾기
    print("\n" + "=" * 80)
    print("📋 기본사항 다이얼로그 분석")
    print("=" * 80)

    basic_dialog = None
    for ctrl in dlg.descendants():
        try:
            if ctrl.class_name() == "#32770":  # Dialog
                text = ctrl.window_text()
                if "기본사항" in text:
                    basic_dialog = ctrl
                    break
        except:
            pass

    if basic_dialog:
        print(f"✓ 기본사항 다이얼로그 발견: HWND=0x{basic_dialog.handle:08X}")

        # 다이얼로그 내부의 컨트롤만
        dialog_children = basic_dialog.descendants()
        print(f"  다이얼로그 내부 컨트롤: {len(dialog_children)}개\n")

        # Edit 컨트롤만
        dialog_edits = []
        for ctrl in dialog_children:
            try:
                if "Edit" in ctrl.class_name():
                    dialog_edits.append(ctrl)
            except:
                pass

        print(f"📝 다이얼로그 내 Edit 컨트롤 {len(dialog_edits)}개:")
        for idx, ctrl in enumerate(dialog_edits):
            try:
                text = ctrl.window_text()
                print(f"  [{idx}] {ctrl.class_name()} - \"{text}\" - HWND=0x{ctrl.handle:08X}")
            except:
                pass
    else:
        print("✗ 기본사항 다이얼로그를 찾을 수 없습니다")

    # Static 텍스트 (라벨) 찾기
    print("\n" + "=" * 80)
    print("🏷️  Static 라벨 분석")
    print("=" * 80)

    static_controls = []
    for ctrl in descendants:
        try:
            if ctrl.class_name() == "Static":
                text = ctrl.window_text()
                if text and any(keyword in text for keyword in ["사번", "성명", "주민", "나이"]):
                    static_controls.append((ctrl, text))
        except:
            pass

    print(f"관련 라벨 {len(static_controls)}개 발견:\n")
    for ctrl, text in static_controls:
        try:
            rect = ctrl.rectangle()
            print(f"  • \"{text}\"")
            print(f"    위치: L={rect.left} T={rect.top} R={rect.right} B={rect.bottom}")
            print(f"    HWND=0x{ctrl.handle:08X}")
            print()
        except:
            pass

    print("=" * 80)
    print("✅ 분석 완료")
    print("=" * 80)


if __name__ == "__main__":
    analyze_basic_tab()
