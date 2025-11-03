from pywinauto import application
import time
import sys

# UTF-8 출력 설정
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

try:
    # win32 백엔드로 연결
    app = application.Application(backend="win32")
    app.connect(title="연말정산추가자료입력")
    dlg = app.window(title="연말정산추가자료입력")

    print("연말정산 윈도우에 연결 성공")

    # 탭 컨트롤 찾기
    tab_control = dlg.child_window(class_name="Afx:TabWnd:330000:8:10003:10", found_index=0)

    if tab_control.exists():
        rect = tab_control.rectangle()
        print(f"탭 컨트롤 위치: {rect}")

        # 탭 영역의 시작점과 크기 계산
        tab_left = rect.left
        tab_top = rect.top
        tab_width = rect.right - rect.left
        tab_height = rect.bottom - rect.top

        print(f"탭 너비: {tab_width}, 높이: {tab_height}")

        # 화면 캡처를 보면 탭들이 나란히 배치되어 있음
        # 소득정보(첫번째), 부양가족(두번째), 신용카드 등(세번째)...
        # 각 탭의 대략적인 너비를 추정 (약 8-10개 탭이 있다고 가정)

        # 부양가족 탭 클릭 - 두 번째 탭 (소득정보 다음)
        # 화면을 보니 부양가족은 대략 두 번째 탭 위치
        estimated_tab_count = 9
        single_tab_width = tab_width / estimated_tab_count

        # 두 번째 탭의 중앙 좌표 (인덱스 1)
        dependency_tab_x = tab_left + int(single_tab_width * 1.5)
        dependency_tab_y = tab_top + int(tab_height / 2)

        print(f"부양가족 탭 클릭 좌표: ({dependency_tab_x}, {dependency_tab_y})")

        # 절대 좌표로 클릭
        from pywinauto import mouse
        mouse.click(coords=(dependency_tab_x, dependency_tab_y))

        print("부양가족 탭 클릭 완료!")
        time.sleep(0.5)

        # 성공 여부 확인을 위해 화면 캡처나 현재 상태 확인
        print("부양가족 탭이 선택되었는지 확인해주세요.")

    else:
        print("탭 컨트롤을 찾을 수 없습니다.")

except Exception as e:
    print(f"오류: {e}")
    import traceback
    traceback.print_exc()
