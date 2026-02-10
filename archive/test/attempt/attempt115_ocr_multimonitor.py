"""
시도 115: 멀티 모니터 대응 OCR

all_screens=True 옵션 또는 윈도우를 프라이머리 모니터로 이동
"""
import win32process
import win32gui
import win32con
import win32api
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
    print("시도 115: 멀티 모니터 대응 OCR")
    print("="*60)

    try:
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
        dlg_hwnd = None
        for hwnd, title in found_dialogs:
            if not title:
                dialog = app.window(handle=hwnd)
                for child in dialog.children():
                    try:
                        text = child.window_text()
                        if "분납적용" in text:
                            installment_dlg = dialog
                            dlg_hwnd = hwnd
                            print(f"✓ 분납적용 다이얼로그: 0x{hwnd:08X}")
                            break
                    except:
                        pass
                if installment_dlg:
                    break

        if not installment_dlg:
            return {"success": False, "message": "분납적용 다이얼로그를 찾지 못했습니다"}

        # 2. 현재 위치 확인
        dlg_rect = installment_dlg.rectangle()
        print(f"\n현재 다이얼로그 위치: ({dlg_rect.left}, {dlg_rect.top})")
        print(f"  크기: {dlg_rect.width()} x {dlg_rect.height()}")

        # 3. 방법 1: all_screens=True로 캡처
        print("\n[방법 1] all_screens=True로 전체 화면 캡처")

        try:
            img_all = ImageGrab.grab(all_screens=True)
            print(f"✓ 전체 화면 캡처 성공: {img_all.size}")

            # 가상 화면에서 다이얼로그 영역 크롭
            # ImageGrab의 좌표계는 가장 왼쪽 위 모니터 기준 (0, 0)
            # 하지만 Win32 좌표계와 다를 수 있음

            # Win32 가상 화면 크기 확인
            sm_xvirtualscreen = win32api.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN
            sm_yvirtualscreen = win32api.GetSystemMetrics(77)  # SM_YVIRTUALSCREEN
            sm_cxvirtualscreen = win32api.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
            sm_cyvirtualscreen = win32api.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN

            print(f"\n가상 화면 정보:")
            print(f"  원점: ({sm_xvirtualscreen}, {sm_yvirtualscreen})")
            print(f"  크기: {sm_cxvirtualscreen} x {sm_cyvirtualscreen}")

            # 좌표 변환: Win32 → ImageGrab
            # ImageGrab는 (0, 0)부터 시작하므로 오프셋 제거
            crop_left = dlg_rect.left - sm_xvirtualscreen
            crop_top = dlg_rect.top - sm_yvirtualscreen
            crop_right = dlg_rect.right - sm_xvirtualscreen
            crop_bottom = dlg_rect.bottom - sm_yvirtualscreen

            print(f"\nImageGrab 좌표계로 변환:")
            print(f"  ({crop_left}, {crop_top}) - ({crop_right}, {crop_bottom})")

            # 크롭
            img_dlg = img_all.crop((crop_left, crop_top, crop_right, crop_bottom))
            print(f"✓ 다이얼로그 영역 크롭: {img_dlg.size}")

            # 저장
            img_dlg.save("test/image/attempt115_dialog_cropped.png")
            print(f"✓ 이미지 저장: test/image/attempt115_dialog_cropped.png")

            # 통계 확인
            hist = img_dlg.histogram()
            avg_r = sum(i * hist[i] for i in range(256)) / sum(hist[:256]) if sum(hist[:256]) > 0 else 0
            avg_g = sum(i * hist[256 + i] for i in range(256)) / sum(hist[256:512]) if sum(hist[256:512]) > 0 else 0
            avg_b = sum(i * hist[512 + i] for i in range(256)) / sum(hist[512:768]) if sum(hist[512:768]) > 0 else 0

            print(f"\n이미지 통계:")
            print(f"  평균 RGB: ({avg_r:.1f}, {avg_g:.1f}, {avg_b:.1f})")

            if avg_r < 30 and avg_g < 30 and avg_b < 30:
                print(f"  → 여전히 검은색 (다른 윈도우에 가려졌거나 최소화)")

                # 방법 2: 윈도우를 앞으로 가져오기
                print("\n[방법 2] 윈도우를 포그라운드로")

                try:
                    win32gui.SetForegroundWindow(dlg_hwnd)
                    time.sleep(0.5)
                    print("✓ SetForegroundWindow 완료")

                    # 다시 캡처
                    img_all2 = ImageGrab.grab(all_screens=True)
                    img_dlg2 = img_all2.crop((crop_left, crop_top, crop_right, crop_bottom))
                    img_dlg2.save("test/image/attempt115_dialog_foreground.png")

                    hist2 = img_dlg2.histogram()
                    avg_r2 = sum(i * hist2[i] for i in range(256)) / sum(hist2[:256]) if sum(hist2[:256]) > 0 else 0
                    avg_g2 = sum(i * hist2[256 + i] for i in range(256)) / sum(hist2[256:512]) if sum(hist2[256:512]) > 0 else 0
                    avg_b2 = sum(i * hist2[512 + i] for i in range(256)) / sum(hist2[512:768]) if sum(hist2[512:768]) > 0 else 0

                    print(f"재캡처 평균 RGB: ({avg_r2:.1f}, {avg_g2:.1f}, {avg_b2:.1f})")

                    if avg_r2 > 30 or avg_g2 > 30 or avg_b2 > 30:
                        print("✓ 재캡처 성공! 내용 있음")
                        img_dlg = img_dlg2  # 새 이미지 사용
                    else:
                        print("✗ 재캡처도 검은색")

                except Exception as e:
                    print(f"✗ SetForegroundWindow 실패: {e}")
            else:
                print(f"  → 내용 있음!")

            # OCR 시도
            print("\n[3단계] OCR 시도")

            try:
                import pytesseract

                # Tesseract 경로
                tesseract_paths = [
                    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                ]

                for path in tesseract_paths:
                    import os
                    if os.path.exists(path):
                        pytesseract.pytesseract.tesseract_cmd = path
                        break

                text = pytesseract.image_to_string(img_dlg, lang='kor+eng')
                print(f"✓ OCR 완료: {len(text)}자")

                if text.strip():
                    print("\n추출된 텍스트 (처음 500자):")
                    print("=" * 60)
                    print(text[:500])
                    print("=" * 60)

                    # 사원코드 파싱
                    lines = text.split('\n')
                    employee_codes = []

                    for line in lines:
                        line = line.strip()
                        if line.isdigit() and 3 <= len(line) <= 10:
                            employee_codes.append(line)

                    if employee_codes:
                        print(f"\n✓✓✓ {len(employee_codes)}개 사원코드 발견:")
                        for code in employee_codes[:10]:
                            print(f"  - {code}")

                        return {
                            "success": True,
                            "message": f"OCR 성공! {len(employee_codes)}개 사원코드 추출"
                        }
                    else:
                        return {
                            "success": False,
                            "message": f"OCR 완료 ({len(text)}자) 하지만 사원코드 패턴 없음"
                        }

            except ImportError:
                return {"success": True, "message": "이미지 저장 성공. pytesseract 설치 필요"}
            except Exception as e:
                return {"success": True, "message": f"이미지 저장 성공. OCR 오류: {e}"}

        except TypeError as e:
            # all_screens 파라미터 지원 안 함
            print(f"✗ all_screens 지원 안 함: {e}")
            return {"success": False, "message": "PIL 버전이 all_screens를 지원하지 않음. PIL 업그레이드 필요"}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
