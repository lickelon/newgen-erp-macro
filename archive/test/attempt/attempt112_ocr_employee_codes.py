"""
시도 112: OCR로 사원코드 읽기

PrintWindow로 스프레드 캡처 후 pytesseract로 사원코드 추출
"""
import win32process
import win32gui
import win32con
import win32api
import win32ui
import time
from pywinauto import Application
from ctypes import windll
from PIL import Image


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 112: OCR로 사원코드 읽기")
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
            print("⚠️  Tesseract 자동 감지 실패")
            print("   수동으로 설정하거나 https://github.com/UB-Mannheim/tesseract/wiki 에서 설치")

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

        left_hwnd = left_spread.handle
        rect = left_spread.rectangle()
        print(f"✓ 왼쪽 스프레드: 0x{left_hwnd:08X}")
        print(f"  위치: ({rect.left}, {rect.top}) - ({rect.right}, {rect.bottom})")
        print(f"  크기: {rect.width()} x {rect.height()}")

        # 3. 스프레드 캡처 함수
        def capture_window(hwnd):
            """PrintWindow로 윈도우 캡처"""
            # 윈도우 DC
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            # 윈도우 크기
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top

            # 비트맵 생성
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            # PrintWindow로 캡처
            result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

            if result == 0:
                print("  ✗ PrintWindow 실패")
                mfcDC.DeleteDC()
                saveDC.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwndDC)
                win32gui.DeleteObject(saveBitMap.GetHandle())
                return None

            # 비트맵 → PIL Image
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)

            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )

            # 리소스 해제
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)

            return img

        # 4. 스프레드 전체 캡처
        print("\n[3단계] 스프레드 캡처")

        img = capture_window(left_hwnd)
        if img is None:
            return {"success": False, "message": "PrintWindow 캡처 실패"}

        print(f"✓ 캡처 성공: {img.size}")

        # 이미지 저장
        img_path = "test/image/attempt112_spread.png"
        img.save(img_path)
        print(f"✓ 이미지 저장: {img_path}")

        # 5. OCR 실행
        print("\n[4단계] OCR로 텍스트 추출")

        try:
            # 한글 + 영어 + 숫자
            text = pytesseract.image_to_string(img, lang='kor+eng')
            print(f"✓ OCR 완료: {len(text)}자")
            print("\n추출된 텍스트:")
            print("=" * 60)
            print(text)
            print("=" * 60)

            # 6. 사원코드 파싱 (숫자 패턴)
            print("\n[5단계] 사원코드 파싱")

            import re

            # 숫자만 있는 라인 찾기 (사원코드로 추정)
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

                # 모든 숫자 포함 라인 출력 (디버깅)
                print("\n숫자가 포함된 모든 라인:")
                for line in lines:
                    if any(c.isdigit() for c in line):
                        print(f"  '{line.strip()}'")

                return {
                    "success": False,
                    "message": "OCR 성공했으나 사원코드 패턴 파싱 실패"
                }

        except Exception as e:
            print(f"✗ OCR 실패: {e}")
            return {
                "success": False,
                "message": f"OCR 오류: {e}"
            }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
