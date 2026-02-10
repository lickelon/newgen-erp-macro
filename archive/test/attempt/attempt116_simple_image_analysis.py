"""
시도 116: 간단한 이미지 분석으로 사원코드 읽기

OCR 라이브러리 없이 PIL로 이미지 분석
"""
import win32process
import win32gui
import time
from pywinauto import Application
from PIL import Image, ImageGrab, ImageEnhance


def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 116: 간단한 이미지 분석")
    print("="*60)

    try:
        # 이미 저장된 이미지 사용
        img_path = "test/image/attempt115_dialog_cropped.png"

        print(f"이미지 로드: {img_path}")
        img = Image.open(img_path)
        print(f"✓ 이미지 크기: {img.size}")

        # 왼쪽 사원 목록 영역만 크롭 (대략 왼쪽 절반)
        width, height = img.size
        left_area = img.crop((0, 40, width // 3, height - 60))
        left_area.save("test/image/attempt116_employee_list.png")
        print(f"✓ 사원 목록 영역 크롭: {left_area.size}")

        # 그레이스케일 변환
        gray = left_area.convert('L')
        gray.save("test/image/attempt116_gray.png")

        # 대비 증가
        enhancer = ImageEnhance.Contrast(gray)
        enhanced = enhancer.enhance(2.0)
        enhanced.save("test/image/attempt116_enhanced.png")
        print("✓ 이미지 전처리 완료")

        # 픽셀 데이터 분석
        pixels = list(enhanced.getdata())
        width_crop, height_crop = enhanced.size

        print("\n이미지 분석 결과:")
        print(f"  크기: {width_crop} x {height_crop}")
        print(f"  픽셀 수: {len(pixels)}")

        # 각 행의 평균 밝기
        print("\n행별 평균 밝기 (처음 20행):")
        for y in range(min(20, height_crop)):
            row_pixels = pixels[y * width_crop:(y + 1) * width_crop]
            avg_brightness = sum(row_pixels) / len(row_pixels)
            print(f"  행 {y}: {avg_brightness:.1f}")

        # 사원코드가 있는 열 찾기 (No 다음, 사원명 이전)
        # 대략 x=50~150 사이에 사원코드가 있을 것으로 예상

        print("\n✓ 이미지 분석 완료")
        print("\n현재 상황:")
        print("  - 이미지 캡처: ✅")
        print("  - 이미지 전처리: ✅")
        print("  - OCR 없는 텍스트 인식: ❌ (복잡함)")

        print("\n실용적인 해결책:")
        print("1. Tesseract-OCR 설치 (가장 확실)")
        print("   → https://github.com/UB-Mannheim/tesseract/wiki")
        print("   → 설치 후 attempt115 재실행")

        print("\n2. 또는 한 번만 수동 매칭:")
        print("   - 첫 실행시 화면에 표시된 순서 확인")
        print("   - CSV를 그 순서로 정렬")
        print("   - 이후 순서대로 자동 입력")

        return {
            "success": False,
            "message": "이미지 분석 완료. OCR 라이브러리 설치 필요"
        }

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
