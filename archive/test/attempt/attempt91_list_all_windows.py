"""
시도 91: 모든 열려있는 창 나열하여 분납적용 창 찾기
"""
from pywinauto import findwindows, Desktop

def run(dlg, capture_func):
    """
    Args:
        dlg: 사용 안 함 (None 가능)
        capture_func: 사용 안 함

    Returns:
        dict: {"success": bool, "message": str}
    """
    print("\n" + "="*60)
    print("시도 91: 모든 열려있는 창 나열")
    print("="*60)

    try:
        print("\n[방법 1] findwindows.find_elements() 사용")
        print("-" * 60)

        windows = findwindows.find_elements()
        count = 0
        found_installment = False

        for w in windows:
            try:
                title = w.name if w.name else ""
                class_name = w.class_name if w.class_name else ""

                if title:  # 제목이 있는 것만
                    count += 1
                    print(f"{count}. '{title}'")
                    print(f"   클래스: {class_name}")

                    # 분납 관련 키워드 체크
                    if "분납" in title:
                        print(f"   >>> ⭐ 분납 관련 창 발견!")
                        found_installment = True
            except:
                pass

        print(f"\n총 {count}개 창 확인")

        if not found_installment:
            print("\n[방법 2] Desktop backend 사용")
            print("-" * 60)

            try:
                desktop = Desktop(backend="win32")
                wins = desktop.windows()

                print(f"Desktop에서 {len(wins)}개 창 발견")

                for i, win in enumerate(wins[:50]):  # 최대 50개
                    try:
                        title = win.window_text()
                        if title:
                            print(f"{i+1}. '{title}'")
                            if "분납" in title:
                                print(f"   >>> ⭐ 분납 관련 창 발견!")
                                found_installment = True
                    except:
                        pass
            except Exception as e:
                print(f"Desktop 방법 실패: {e}")

        if found_installment:
            return {"success": True, "message": "분납 관련 창을 찾았습니다"}
        else:
            return {"success": False, "message": "분납 관련 창을 찾지 못했습니다. 창이 열려있는지 확인하세요."}

    except Exception as e:
        import traceback
        return {"success": False, "message": f"오류: {e}\n{traceback.format_exc()}"}
