"""
시도 113: ImageGrab으로 OCR

PrintWindow 대신 PIL.ImageGrab으로 화면 영역 캡처
"""
import win32process
import win32gui
import time
from pywinauto import Application
from PIL import ImageGrab


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 113: ImageGrab으로 OCR")
    print("="*60)

    try:
        # pytesseract 확인
        try:
            import pytesseract
            print("✓ pytesseract 로드 성공")
        except ImportError:
            return {
                "success": False,
                "message": "pytesseract 설치 필요: pip install pytesseract"
            }

        # Tesseract 경로 설정
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]

        tesseract_found = False
        for path in tesseract_paths:
            try:
                import os
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    tesseract_found = True
                    print(f"✓ Tesseract 경로: {path}")
                    break
            except:
                pass

        if not tesseract_found:
            print("⚠️  Tesseract 바이너리 없음")
            print("   계속 시도하지만 OCR은 실패할 수 있음")

        # 1. 분납적용 다이얼로그 찾기
        app = Application(backend="win32")
        app.connect(title="급여자료입력")
        main_window = app.window(title="급여자료입력")
        _, process_id = win32process.GetWindowThreadProcessId(main_window.handle)

        print(f"\n[1단계] 급여자료입력 연결 (PID: {process_id})")

        found_dialogs = []

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == process_id:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "#32770":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, found_dialogs)

        installment_dlg = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 2. 스프레드 찾기
        print("\n[2단계] 스프레드 찾기")

        spreads = []
        for child in installment_dlg.children():
            try:
                if child.class_name() == "fpUSpread80":
                    spreads.append(child)
            except:
                pass

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]

        # 스프레드의 화면 좌표
        rect = left_spread.rectangle()
        print(f"✓ 왼쪽 스프레드: 0x{left_spread.handle:08X}")
        print(f"  화면 좌표: ({rect.left}, {rect.top}) - ({rect.right}, {rect.bottom})")
        print(f"  크기: {rect.width()} x {rect.height()}")

        # 3. ImageGrab으로 화면 영역 캡처
        print("\n[3단계] 화면 영역 캡처 (ImageGrab)")

        # 화면 좌표로 캡처
        bbox = (rect.left, rect.top, rect.right, rect.bottom)
        img = ImageGrab.grab(bbox)

        print(f"✓ 캡처 성공: {img.size}")

        # 이미지 저장
        img_path = "test/image/attempt113_spread.png"
        img.save(img_path)
        print(f"✓ 이미지 저장: {img_path}")

        # 4. OCR 실행
        print("\n[4단계] OCR로 텍스트 추출")

        try:
            # 한글 + 영어 + 숫자
            text = pytesseract.image_to_string(img, lang='kor+eng')
            print(f"✓ OCR 완료: {len(text)}자")
            print("\n추출된 텍스트:")
            print("=" * 60)
            print(text)
            print("=" * 60)

            # 5. 사원코드 파싱
            print("\n[5단계] 사원코드 파싱")

            import re

            lines = text.split('\n')
            employee_codes = []

            for line in lines:
                line = line.strip()
                # 3~10자리 숫자 패턴
                if line.isdigit() and 3 <= len(line) <= 10:
                    employee_codes.append(line)
                    print(f"  발견: '{line}'")

            if employee_codes:
                print(f"\n✓✓✓ {len(employee_codes)}개 사원코드 추출 성공!")
                return {
                    "success": True,
                    "message": f"OCR로 {len(employee_codes)}개 사원코드 추출: {employee_codes[:5]}"
                }
            else:
                print("\n✗ 사원코드 패턴을 찾지 못함")

                # 숫자 포함 라인 출력
                print("\n숫자가 포함된 모든 라인:")
                for line in lines:
                    if any(c.isdigit() for c in line):
                        print(f"  '{line.strip()}'")

                return {
                    "success": False,
                    "message": f"OCR 성공했으나 사원코드 파싱 실패. 이미지: {img_path}"
                }

        except Exception as e:
            print(f"✗ OCR 실패: {e}")
            return {
                "success": False,
                "message": f"OCR 오류: {e}. 이미지는 저장됨: {img_path}"
            }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
