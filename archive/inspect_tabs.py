from pywinauto import application
import sys

# UTF-8 출력 설정
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

try:
    # win32 백엔드로 연결
    app = application.Application(backend="win32")
    app.connect(title="연말정산추가자료입력")
    dlg = app.window(title="연말정산추가자료입력")

    print("=== 탭 컨트롤 상세 분석 ===\n")

    # 모든 TabWnd 찾기
    tabs = dlg.children(class_name_re=".*Tab.*")
    print(f"Tab 관련 컨트롤 {len(tabs)}개 발견\n")

    for i, tab in enumerate(tabs):
        print(f"[탭 컨트롤 {i}]")
        print(f"  클래스: {tab.class_name()}")
        print(f"  위치: {tab.rectangle()}")
        print(f"  보이는 상태: {tab.is_visible()}")
        print(f"  활성 상태: {tab.is_enabled()}")

        # 자식 요소 확인
        children = tab.children()
        print(f"  자식 요소 수: {len(children)}")

        for j, child in enumerate(children[:5]):  # 처음 5개만
            try:
                print(f"    자식[{j}]: {child.class_name()}, 제목: {child.window_text()}")
            except:
                print(f"    자식[{j}]: 정보 읽기 실패")
        print()

    # Static 텍스트로 탭 이름 찾기 시도
    print("\n=== Static 텍스트 검색 (탭 이름) ===")
    statics = dlg.children(class_name="Static")
    for i, static in enumerate(statics[:20]):  # 처음 20개만
        try:
            text = static.window_text()
            if text and any(keyword in text for keyword in ["소득", "부양", "신용", "의료", "기부"]):
                print(f"[{i}] '{text}' at {static.rectangle()}")
        except:
            pass

    # 키보드 단축키 테스트
    print("\n=== 키보드 방식 시도 ===")
    print("Ctrl+Tab으로 다음 탭으로 이동 시도...")
    dlg.set_focus()
    dlg.type_keys("^{TAB}")  # Ctrl+Tab
    print("완료. 탭이 변경되었는지 확인해주세요.")

except Exception as e:
    print(f"오류: {e}")
    import traceback
    traceback.print_exc()
