from pywinauto import findwindows
from pywinauto import application
import sys

# UTF-8 출력 설정
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

# 먼저 실행 중인 모든 윈도우 찾기
print("실행 중인 윈도우 목록:")
print("="*50)
windows = findwindows.find_elements()
target_found = False

for w in windows:
    try:
        title = w.name if w.name else "(제목 없음)"
        class_name = w.class_name if w.class_name else "(클래스 없음)"
        print(f"Title: {title}, Class: {class_name}")

        if "사원등록" in title:
            target_found = True
            print(f"  >>> 사원등록 윈도우 발견!")
    except Exception as e:
        print(f"윈도우 정보 출력 중 오류: {e}")

print("\n" + "="*50 + "\n")

if not target_found:
    print("⚠️  '사원등록' 윈도우를 찾을 수 없습니다.")
    print("프로그램이 실행 중인지 확인해주세요.")
    sys.exit(1)

# 사원등록 윈도우에 연결
try:
    print("'사원등록' 윈도우에 연결 중...")

    # win32 백엔드로 시도 (MFC 애플리케이션에 더 적합)
    try:
        app = application.Application(backend="win32")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        print("✓ win32 백엔드로 연결 성공")
    except:
        # win32로 안되면 UIA로 재시도
        print("win32 백엔드 실패, UIA로 재시도...")
        app = application.Application(backend="uia")
        app.connect(title="사원등록")
        dlg = app.window(title="사원등록")
        print("✓ UIA 백엔드로 연결 성공")

    # print("\n컨트롤 식별자 목록:")
    # print("="*50)
    # dlg.print_control_identifiers()
    dlg = dlg.child_window(class_name="AfxFrameOrView90u").child_window(class_name="Afx:TabWnd:cd0000:8:10003:10").child_window(title=" 부양가족명세 ", class_name="#32770")
    # dlg.child_window(title=" 부양가족명세 ", class_name="#32770")
    dlg.print_control_identifiers()

except Exception as e:
    print(f"\n❌ 오류 발생: {e}")
    print("프로그램이 실행 중인지, 정확한 윈도우 제목인지 확인해주세요.")
    sys.exit(1)