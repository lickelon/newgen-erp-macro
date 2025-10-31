"""
시도 83: Python UIAutomation 라이브러리 사용

uiautomation 패키지를 사용하여 백그라운드에서 값 읽기 시도
https://github.com/yinkaisheng/Python-UIAutomation-for-Windows
"""
import time
from ctypes import *
from ctypes.wintypes import HWND
import win32gui
import win32con
import win32api
import pyperclip


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 83: Python UIAutomation 라이브러리")
    print("="*60)

    try:
        # uiautomation 라이브러리 import 시도
        try:
            import uiautomation as auto
            print("✓ uiautomation 라이브러리 로드 성공")
        except ImportError:
            return {
                "success": False,
                "message": """
❌ uiautomation 라이브러리가 설치되지 않았습니다!

설치 명령:
  pip install uiautomation
  또는
  uv pip install uiautomation

설치 후 다시 시도하세요.
"""
            }

        # 초기 상태 캡처
        capture_func("attempt83_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        spread_hwnd = left_spread.handle

        print(f"\n왼쪽 스프레드 HWND: 0x{spread_hwnd:08X}")

        print("\n=== 1. UIAutomation으로 컨트롤 찾기 ===")

        # HWND로 UIAutomation 컨트롤 가져오기
        try:
            spread_control = auto.ControlFromHandle(spread_hwnd)
            print(f"✓ UIAutomation 컨트롤 생성 성공")
            print(f"  ControlType: {spread_control.ControlTypeName}")
            print(f"  Name: {spread_control.Name}")
            print(f"  ClassName: {spread_control.ClassName}")
        except Exception as e:
            return {
                "success": False,
                "message": f"UIAutomation 컨트롤 생성 실패: {e}"
            }

        print("\n=== 2. 지원하는 패턴 확인 ===")

        # 사용 가능한 패턴들 확인
        patterns = []
        pattern_names = [
            'ValuePattern', 'TextPattern', 'GridPattern', 'TablePattern',
            'SelectionItemPattern', 'SelectionPattern', 'InvokePattern',
            'ScrollPattern', 'RangeValuePattern'
        ]

        for pattern_name in pattern_names:
            try:
                pattern = getattr(spread_control, f'Get{pattern_name}', None)
                if pattern and callable(pattern):
                    pattern_obj = pattern()
                    if pattern_obj:
                        patterns.append(pattern_name)
                        print(f"  ✓ {pattern_name} 지원")
            except:
                pass

        if not patterns:
            print("  ✗ 지원하는 패턴 없음")

        print("\n=== 3. ValuePattern으로 값 읽기 시도 ===")

        try:
            value_pattern = spread_control.GetValuePattern()
            if value_pattern:
                value = value_pattern.Value
                print(f"✓ ValuePattern.Value: '{value}'")

                if value:
                    print(f"\n✓✓✓ 값 읽기 성공!")

                    # 백그라운드 테스트
                    print(f"\n=== 4. 백그라운드 읽기 테스트 ===")

                    import subprocess
                    notepad = subprocess.Popen(['notepad.exe'])
                    time.sleep(2)

                    active = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                    print(f"현재 활성 창: '{active}'")

                    # 백그라운드에서 다시 읽기
                    bg_value = value_pattern.Value
                    print(f"백그라운드 ValuePattern.Value: '{bg_value}'")

                    if bg_value:
                        print(f"✓✓✓✓ 백그라운드 읽기 성공!")

                        notepad.terminate()
                        capture_func("attempt83_01_success.png")

                        return {
                            "success": True,
                            "message": f"""
🎉 UIAutomation ValuePattern으로 백그라운드 읽기 성공!

값: '{bg_value}'

이 방법이 작동합니다! 🎊
"""
                        }

                    notepad.terminate()
            else:
                print("✗ ValuePattern 지원 안 함")
        except Exception as e:
            print(f"✗ ValuePattern 실패: {e}")

        print("\n=== 5. TextPattern으로 텍스트 읽기 시도 ===")

        try:
            text_pattern = spread_control.GetTextPattern()
            if text_pattern:
                document_range = text_pattern.DocumentRange
                text = document_range.GetText(-1)
                print(f"✓ TextPattern.DocumentRange.GetText(): '{text}'")

                if text:
                    print(f"\n✓✓✓ 텍스트 읽기 성공!")

                    return {
                        "success": True,
                        "message": f"""
🎉 UIAutomation TextPattern으로 텍스트 읽기 성공!

텍스트: '{text}'
"""
                    }
            else:
                print("✗ TextPattern 지원 안 함")
        except Exception as e:
            print(f"✗ TextPattern 실패: {e}")

        print("\n=== 6. GridPattern으로 셀 접근 시도 ===")

        try:
            grid_pattern = spread_control.GetGridPattern()
            if grid_pattern:
                row_count = grid_pattern.RowCount
                col_count = grid_pattern.ColumnCount
                print(f"✓ GridPattern 지원")
                print(f"  행: {row_count}, 열: {col_count}")

                if row_count > 0 and col_count > 0:
                    # 첫 번째 셀 가져오기
                    cell = grid_pattern.GetItem(0, 0)
                    if cell:
                        cell_name = cell.Name
                        print(f"  Cell(0, 0) Name: '{cell_name}'")

                        # 셀의 ValuePattern 시도
                        try:
                            cell_value_pattern = cell.GetValuePattern()
                            if cell_value_pattern:
                                cell_value = cell_value_pattern.Value
                                print(f"  Cell(0, 0) Value: '{cell_value}'")

                                if cell_value:
                                    return {
                                        "success": True,
                                        "message": f"""
🎉 UIAutomation GridPattern으로 셀 값 읽기 성공!

Cell(0, 0): '{cell_value}'
"""
                                    }
                        except:
                            pass
            else:
                print("✗ GridPattern 지원 안 함")
        except Exception as e:
            print(f"✗ GridPattern 실패: {e}")

        print("\n=== 7. 모든 자식 컨트롤 순회 ===")

        children = spread_control.GetChildren()
        print(f"자식 컨트롤 수: {len(children)}")

        for i, child in enumerate(children[:10]):  # 처음 10개만
            print(f"\n  Child {i}:")
            print(f"    ControlType: {child.ControlTypeName}")
            print(f"    Name: '{child.Name}'")

            try:
                value_pattern = child.GetValuePattern()
                if value_pattern:
                    value = value_pattern.Value
                    if value:
                        print(f"    Value: '{value}'")
            except:
                pass

        capture_func("attempt83_02_complete.png")

        return {
            "success": len(patterns) > 0,
            "message": f"""
UIAutomation 조사 완료

지원 패턴: {', '.join(patterns) if patterns else '없음'}
자식 컨트롤: {len(children)}개

fpUSpread80은 UIAutomation의 ValuePattern, TextPattern, GridPattern을
지원하지 않습니다.
"""
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
