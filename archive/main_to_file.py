from pywinauto import findwindows
from pywinauto import application
import sys

# 파일에 UTF-8로 저장
output_file = "result.txt"

with open(output_file, 'w', encoding='utf-8') as f:
    # 먼저 실행 중인 모든 윈도우 찾기
    f.write("실행 중인 윈도우 목록:\n")
    f.write("="*50 + "\n")
    windows = findwindows.find_elements()
    target_found = False

    for w in windows:
        try:
            title = w.name if w.name else "(제목 없음)"
            class_name = w.class_name if w.class_name else "(클래스 없음)"
            f.write(f"Title: {title}, Class: {class_name}\n")

            if "사원등록" in title:
                target_found = True
                f.write(f"  >>> 사원등록 윈도우 발견!\n")
        except Exception as e:
            f.write(f"윈도우 정보 출력 중 오류: {e}\n")

    f.write("\n" + "="*50 + "\n\n")

    if not target_found:
        f.write("⚠️  '사원등록' 윈도우를 찾을 수 없습니다.\n")
        f.write("프로그램이 실행 중인지 확인해주세요.\n")
        sys.exit(1)

    # 사원등록 윈도우에 연결
    try:
        f.write("'사원등록' 윈도우에 연결 중...\n")

        # win32 백엔드로 시도 (MFC 애플리케이션에 더 적합)
        try:
            app = application.Application(backend="win32")
            app.connect(title="사원등록")
            dlg = app.window(title="사원등록")
            f.write("✓ win32 백엔드로 연결 성공\n")
        except:
            # win32로 안되면 UIA로 재시도
            f.write("win32 백엔드 실패, UIA로 재시도...\n")
            app = application.Application(backend="uia")
            app.connect(title="사원등록")
            dlg = app.window(title="사원등록")
            f.write("✓ UIA 백엔드로 연결 성공\n")

        f.write("\n컨트롤 식별자 목록:\n")
        f.write("="*50 + "\n")

        # print_control_identifiers의 출력을 파일로 리다이렉트
        import io
        from contextlib import redirect_stdout

        string_buffer = io.StringIO()
        with redirect_stdout(string_buffer):
            dlg.print_control_identifiers()

        f.write(string_buffer.getvalue())

    except Exception as e:
        f.write(f"\n❌ 오류 발생: {e}\n")
        f.write("프로그램이 실행 중인지, 정확한 윈도우 제목인지 확인해주세요.\n")
        sys.exit(1)

print(f"결과가 {output_file}에 UTF-8 인코딩으로 저장되었습니다.")
