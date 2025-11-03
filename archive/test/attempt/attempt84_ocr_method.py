"""
시도 84: OCR (광학 문자 인식) 방식

백그라운드 윈도우를 캡처하여 OCR로 텍스트 읽기
pytesseract + PIL 사용
"""
import time
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api
import win32ui
from PIL import Image


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 84: OCR (광학 문자 인식) 방식")
    print("="*60)

    try:
        # pytesseract import 시도
        try:
            import pytesseract
            print("✓ pytesseract 라이브러리 로드 성공")
        except ImportError:
            return {
                "success": False,
                "message": """
❌ pytesseract 라이브러리가 설치되지 않았습니다!

설치 명령:
  pip install pytesseract pillow
  또는
  uv pip install pytesseract pillow

그리고 Tesseract-OCR 바이너리도 설치해야 합니다:
  https://github.com/UB-Mannheim/tesseract/wiki

설치 후 pytesseract.pytesseract.tesseract_cmd를 설정하세요.
"""
            }

        # Tesseract 경로 설정 (일반적인 설치 경로)
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
                    print(f"✓ Tesseract 경로 설정: {path}")
                    break
            except:
                pass

        if not tesseract_found:
            print("⚠️  Tesseract 경로를 자동으로 찾지 못했습니다.")
            print("   pytesseract.pytesseract.tesseract_cmd를 수동으로 설정하세요.")

        # 초기 상태 캡처
        capture_func("attempt84_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        spread_hwnd = left_spread.handle

        print(f"\n왼쪽 스프레드 HWND: 0x{spread_hwnd:08X}")

        print("\n=== 1. 백그라운드 윈도우 캡처 ===")

        def capture_window_region(hwnd, left, top, right, bottom):
            """
            윈도우의 특정 영역을 백그라운드에서 캡처
            PrintWindow API 사용
            """
            # 윈도우 DC 가져오기
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            # 비트맵 생성
            width = right - left
            height = bottom - top

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            # PrintWindow로 윈도우 내용 캡처 (백그라운드 가능)
            result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

            if result == 0:
                print("  ✗ PrintWindow 실패")
                mfcDC.DeleteDC()
                saveDC.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwndDC)
                win32gui.DeleteObject(saveBitMap.GetHandle())
                return None

            # 비트맵을 PIL Image로 변환
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)

            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )

            # 영역 크롭
            img = img.crop((0, 0, width, height))

            # 리소스 해제
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)

            return img

        # 스프레드 영역 캡처 (첫 번째 셀 부분)
        rect = left_spread.rectangle()
        print(f"  스프레드 영역: ({rect.left}, {rect.top}) - ({rect.right}, {rect.bottom})")

        # 첫 번째 셀 부분만 캡처 (대략적인 좌표)
        cell_left = 50
        cell_top = 30
        cell_right = 200
        cell_bottom = 60

        print(f"  캡처 영역: ({cell_left}, {cell_top}) - ({cell_right}, {cell_bottom})")

        img = capture_window_region(spread_hwnd, cell_left, cell_top, cell_right, cell_bottom)

        if img is None:
            return {
                "success": False,
                "message": "PrintWindow로 캡처 실패"
            }

        print(f"✓ 이미지 캡처 성공: {img.size}")

        # 캡처된 이미지 저장
        img_path = "test/image/attempt84_captured_cell.png"
        img.save(img_path)
        print(f"✓ 이미지 저장: {img_path}")

        print("\n=== 2. OCR로 텍스트 추출 ===")

        try:
            # OCR 실행
            text = pytesseract.image_to_string(img, lang='eng+kor')
            print(f"✓ OCR 결과: '{text.strip()}'")

            if text.strip():
                # 백그라운드 테스트
                print(f"\n=== 3. 백그라운드 OCR 테스트 ===")

                import subprocess
                notepad = subprocess.Popen(['notepad.exe'])
                time.sleep(2)

                active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                print(f"현재 활성 창: '{active}'")

                # 백그라운드에서 다시 캡처 및 OCR
                img2 = capture_window_region(spread_hwnd, cell_left, cell_top, cell_right, cell_bottom)

                if img2:
                    text2 = pytesseract.image_to_string(img2, lang='eng+kor')
                    print(f"백그라운드 OCR 결과: '{text2.strip()}'")

                    if text2.strip():
                        print(f"\n✓✓✓✓ 백그라운드 OCR 성공!")

                        notepad.terminate()
                        capture_func("attempt84_01_success.png")

                        return {
                            "success": True,
                            "message": f"""
🎉 OCR 방식으로 백그라운드 읽기 성공!

포그라운드: '{text.strip()}'
백그라운드: '{text2.strip()}'

이 방법이 작동합니다! 🎊

단점:
- OCR 정확도에 의존
- 느린 속도
- Tesseract 설치 필요

장점:
- 백그라운드에서 작동
- 모든 컨트롤에 적용 가능
"""
                        }

                notepad.terminate()

                return {
                    "success": True,
                    "message": f"""
OCR로 텍스트 읽기 성공 (포그라운드)

텍스트: '{text.strip()}'

백그라운드 캡처는 실패했지만, 포그라운드에서는 작동합니다.
"""
                }

        except Exception as e:
            print(f"✗ OCR 실패: {e}")

            return {
                "success": False,
                "message": f"""
OCR 실패: {e}

Tesseract-OCR이 제대로 설치되었는지 확인하세요:
https://github.com/UB-Mannheim/tesseract/wiki
"""
            }

        capture_func("attempt84_02_complete.png")

        return {
            "success": False,
            "message": "OCR로 텍스트를 읽지 못했습니다."
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
