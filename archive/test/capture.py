"""스크린샷 캡처 유틸리티"""
import os
import win32gui
import mss
from PIL import Image


def capture_window(hwnd, filename):
    """
    윈도우 핸들로 스크린샷 캡처

    Args:
        hwnd: 윈도우 핸들
        filename: 저장할 파일명

    Returns:
        PIL Image 객체
    """
    # test/image/ 폴더에 저장
    os.makedirs("test/image", exist_ok=True)
    filepath = os.path.join("test", "image", filename)

    # 윈도우 영역 가져오기
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)

    # mss로 스크린샷 캡처
    with mss.mss() as sct:
        monitor = {"top": top, "left": left, "width": right - left, "height": bottom - top}
        screenshot = sct.grab(monitor)
        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
        img.save(filepath)
        return img
