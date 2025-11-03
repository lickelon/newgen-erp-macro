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

    # 메인 윈도우 가져오기
    dlg = app.window(title="연말정산추가자료입력")

    print("연말정산 윈도우에 연결 성공")

    # 탭 컨트롤 찾기
    # Afx:TabWnd 클래스의 탭 컨트롤 찾기
    tabs = dlg.children(class_name="Afx:TabWnd:330000:8:10003:10")

    if tabs:
        print(f"탭 컨트롤 {len(tabs)}개 발견")

        # 첫 번째 탭 컨트롤 (메인 탭)
        main_tab = tabs[0]

        # 탭의 rectangle 정보 가져오기
        rect = main_tab.rectangle()
        print(f"탭 위치: {rect}")

        # 부양가족 탭의 대략적인 위치 계산
        # 화면 캡처를 보면 소득정보 다음에 부양가족이 있음
        # 두 번째 탭을 클릭하기 위해 탭 영역의 왼쪽에서 두 번째 위치 클릭

        tab_width = (rect.right - rect.left) / 10  # 대략 10개 탭이 있다고 가정
        second_tab_x = rect.left + int(tab_width * 1.5)  # 두 번째 탭 중앙
        tab_y = rect.top + 10  # 탭 상단에서 약간 아래

        print(f"부양가족 탭 클릭 예정 위치: ({second_tab_x}, {tab_y})")

        # 탭 클릭
        main_tab.click_input(coords=(second_tab_x - rect.left, tab_y - rect.top))

        print("부양가족 탭 클릭 완료")
        time.sleep(1)

    else:
        print("탭 컨트롤을 찾을 수 없습니다")

        # 대체 방법: 화면 좌표로 직접 클릭
        print("화면 좌표로 직접 클릭 시도...")
        # 화면 캡처에서 부양가족 탭의 대략적인 위치로 클릭
        # 실제 좌표는 화면마다 다를 수 있으므로 조정 필요
        dlg.set_focus()

except Exception as e:
    print(f"오류 발생: {e}")
    import traceback
    traceback.print_exc()
