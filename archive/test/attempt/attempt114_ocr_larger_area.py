"""
시도 114: 더 큰 영역 캡처

스프레드의 실제 내용이 보이는 더 큰 영역을 캡처
"""
import win32process
import win32gui
import time
from pywinauto import Application
from PIL import ImageGrab, ImageDraw, ImageFont


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 114: 더 큰 영역 캡처")
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

        # 2. 다이얼로그 전체 영역 확인
        dlg_rect = installment_dlg.rectangle()
        print(f"\n다이얼로그 전체: ({dlg_rect.left}, {dlg_rect.top}) - ({dlg_rect.right}, {dlg_rect.bottom})")
        print(f"  크기: {dlg_rect.width()} x {dlg_rect.height()}")

        # 3. 스프레드 찾기
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

        spread_rect = left_spread.rectangle()
        print(f"✓ 왼쪽 스프레드: 0x{left_spread.handle:08X}")
        print(f"  화면 좌표: ({spread_rect.left}, {spread_rect.top}) - ({spread_rect.right}, {spread_rect.bottom})")
        print(f"  크기: {spread_rect.width()} x {spread_rect.height()}")

        # 4. 여러 크기로 캡처 시도
        print("\n[3단계] 여러 크기로 캡처")

        captures = []

        # 4-1. 스프레드 영역 그대로
        print("\n[캡처 1] 스프레드 영역 그대로")
        bbox1 = (spread_rect.left, spread_rect.top, spread_rect.right, spread_rect.bottom)
        img1 = ImageGrab.grab(bbox1)
        print(f"  크기: {img1.size}")
        img1.save("test/image/attempt114_spread_original.png")
        captures.append(("원본", img1))

        # 4-2. 스프레드 + 위아래 여유 (100픽셀)
        print("\n[캡처 2] 스프레드 + 위아래 100픽셀")
        bbox2 = (spread_rect.left, spread_rect.top - 100, spread_rect.right, spread_rect.bottom + 100)
        img2 = ImageGrab.grab(bbox2)
        print(f"  크기: {img2.size}")
        img2.save("test/image/attempt114_spread_padded.png")
        captures.append(("100px 여유", img2))

        # 4-3. 다이얼로그 왼쪽 절반
        print("\n[캡처 3] 다이얼로그 왼쪽 절반")
        mid_x = (dlg_rect.left + dlg_rect.right) // 2
        bbox3 = (dlg_rect.left, dlg_rect.top, mid_x, dlg_rect.bottom)
        img3 = ImageGrab.grab(bbox3)
        print(f"  크기: {img3.size}")
        img3.save("test/image/attempt114_dialog_left.png")
        captures.append(("다이얼로그 왼쪽", img3))

        # 4-4. 다이얼로그 전체
        print("\n[캡처 4] 다이얼로그 전체")
        bbox4 = (dlg_rect.left, dlg_rect.top, dlg_rect.right, dlg_rect.bottom)
        img4 = ImageGrab.grab(bbox4)
        print(f"  크기: {img4.size}")
        img4.save("test/image/attempt114_dialog_full.png")
        captures.append(("다이얼로그 전체", img4))

        print(f"\n✓ {len(captures)}개 이미지 저장 완료")

        # 5. 각 이미지의 통계 확인
        print("\n[4단계] 이미지 통계")

        for name, img in captures:
            # 픽셀 히스토그램으로 이미지가 비어있는지 확인
            hist = img.histogram()
            # RGB 각 채널의 평균
            avg_r = sum(i * hist[i] for i in range(256)) / sum(hist[:256])
            avg_g = sum(i * hist[256 + i] for i in range(256)) / sum(hist[256:512])
            avg_b = sum(i * hist[512 + i] for i in range(256)) / sum(hist[512:768])

            print(f"\n  [{name}]")
            print(f"    크기: {img.size}")
            print(f"    평균 RGB: ({avg_r:.1f}, {avg_g:.1f}, {avg_b:.1f})")

            # 검은색에 가까우면 비어있을 가능성
            if avg_r < 30 and avg_g < 30 and avg_b < 30:
                print(f"    → 거의 검은색 (비어있음?)")
            elif avg_r > 200 and avg_g > 200 and avg_b > 200:
                print(f"    → 거의 흰색 (배경?)")
            else:
                print(f"    → 내용 있음 (가능성 높음)")

        # 6. 가장 유망한 이미지 선택
        print("\n[5단계] OCR 시도")

        # pytesseract 확인
        try:
            import pytesseract
            print("✓ pytesseract 로드 성공")

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
                print("⚠️  Tesseract 없음 - OCR 스킵")
                return {
                    "success": True,
                    "message": f"{len(captures)}개 이미지 저장. Tesseract 설치 후 OCR 가능"
                }

            # 다이얼로그 왼쪽 이미지로 OCR
            print("\n다이얼로그 왼쪽 이미지로 OCR 시도:")
            text = pytesseract.image_to_string(img3, lang='kor+eng')
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

            return {
                "success": True,
                "message": f"이미지 저장 성공. OCR 결과: {len(text)}자 (사원코드 파싱 필요)"
            }

        except ImportError:
            return {
                "success": True,
                "message": f"{len(captures)}개 이미지 저장 완료. pytesseract 설치 필요"
            }
        except Exception as e:
            return {
                "success": True,
                "message": f"이미지 저장 성공. OCR 오류: {e}"
            }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
