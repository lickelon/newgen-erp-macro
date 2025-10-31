"""
시도 68: 포괄적 레지스트리 검색

Fpspr80.ocx 파일 경로 검색, ProgID 직접 검색 등
더 광범위한 방법으로 FarPoint Spread 컴포넌트 찾기
"""
import time
import subprocess
from ctypes import *
from ctypes.wintypes import HWND, LONG, DWORD
import winreg
import os


def run(dlg, capture_func):
    print("\n" + "="*60)
    print("시도 68: 포괄적 레지스트리 검색")
    print("="*60)

    try:
        # 초기 상태 캡처
        capture_func("attempt68_00_initial.png")

        # 왼쪽 스프레드 찾기
        spreads = dlg.children(class_name="fpUSpread80")
        if not spreads:
            return {"success": False, "message": "fpUSpread80을 찾을 수 없음"}

        spreads.sort(key=lambda s: s.rectangle().left)
        left_spread = spreads[0]
        hwnd = left_spread.handle

        print(f"왼쪽 스프레드 HWND: 0x{hwnd:08X}")

        # 포커스 설정
        left_spread.set_focus()
        time.sleep(0.5)

        # 참조 값 확인
        import pyperclip
        pyperclip.copy("BEFORE")
        left_spread.type_keys("^c", pause=0.05)
        time.sleep(0.3)
        reference_value = pyperclip.paste()
        print(f"참조 값 (복사): '{reference_value}'")

        print("\n=== 1. Fpspr80.ocx 파일 경로 검색 ===")

        ocx_paths = []
        try:
            # CLSID 하위의 InprocServer32 검색
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID") as clsid_key:
                i = 0
                while i < 10000:  # 10000개까지 검색
                    try:
                        clsid_name = winreg.EnumKey(clsid_key, i)
                        try:
                            with winreg.OpenKey(clsid_key, f"{clsid_name}\\InprocServer32") as server_key:
                                path = winreg.QueryValue(server_key, "")
                                if path and "fpspr" in path.lower():
                                    print(f"  발견: {clsid_name}")
                                    print(f"    경로: {path}")

                                    # ProgID 찾기
                                    progid = None
                                    try:
                                        with winreg.OpenKey(clsid_key, f"{clsid_name}\\ProgID") as progid_key:
                                            progid = winreg.QueryValue(progid_key, "")
                                            print(f"    ProgID: {progid}")
                                    except:
                                        pass

                                    ocx_paths.append({
                                        'clsid': clsid_name,
                                        'path': path,
                                        'progid': progid
                                    })
                        except:
                            pass
                        i += 1
                    except OSError:
                        break
        except Exception as e:
            print(f"검색 오류: {e}")

        if ocx_paths:
            print(f"\n✓ {len(ocx_paths)}개의 Fpspr 관련 컴포넌트 발견")
        else:
            print("\n✗ Fpspr80.ocx를 찾을 수 없음")

        print("\n=== 2. ProgID 직접 검색 ===")

        found_progids = []
        try:
            # HKEY_CLASSES_ROOT 직접 검색
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as root:
                i = 0
                while i < 15000:  # 15000개까지 검색
                    try:
                        key_name = winreg.EnumKey(root, i)

                        # FP, Spread, vaFP 등으로 시작하는 것 찾기
                        if any(key_name.lower().startswith(prefix) for prefix in ['fp', 'spread', 'vafp', 'farpoint']):
                            try:
                                with winreg.OpenKey(root, f"{key_name}\\CLSID") as clsid_subkey:
                                    clsid = winreg.QueryValue(clsid_subkey, "")
                                    print(f"  발견: {key_name}")
                                    print(f"    CLSID: {clsid}")
                                    found_progids.append({
                                        'progid': key_name,
                                        'clsid': clsid
                                    })
                            except:
                                pass

                        i += 1
                    except OSError:
                        break
        except Exception as e:
            print(f"검색 오류: {e}")

        if found_progids:
            print(f"\n✓ {len(found_progids)}개의 ProgID 발견")
        else:
            print("\n✗ 관련 ProgID를 찾을 수 없음")

        print("\n=== 3. TypeLib 검색 ===")

        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "TypeLib") as typelib_key:
                i = 0
                while i < 5000:
                    try:
                        typelib_guid = winreg.EnumKey(typelib_key, i)
                        try:
                            # 버전 검색
                            with winreg.OpenKey(typelib_key, typelib_guid) as guid_key:
                                j = 0
                                while j < 10:
                                    try:
                                        version = winreg.EnumKey(guid_key, j)
                                        try:
                                            desc = winreg.QueryValue(guid_key, version)
                                            if desc and ("spread" in desc.lower() or "farpoint" in desc.lower() or "fpspread" in desc.lower()):
                                                print(f"  발견: {desc}")
                                                print(f"    GUID: {typelib_guid}")
                                                print(f"    버전: {version}")
                                        except:
                                            pass
                                        j += 1
                                    except OSError:
                                        break
                        except:
                            pass
                        i += 1
                    except OSError:
                        break
        except Exception as e:
            print(f"TypeLib 검색 오류: {e}")

        print("\n=== 4. Fpspr80.ocx 파일 시스템 검색 ===")

        search_paths = [
            r"C:\Windows\System32",
            r"C:\Windows\SysWOW64",
            os.path.expandvars(r"%ProgramFiles%"),
            os.path.expandvars(r"%ProgramFiles(x86)%"),
        ]

        for search_path in search_paths:
            if os.path.exists(search_path):
                print(f"\n검색: {search_path}")
                for root, dirs, files in os.walk(search_path):
                    for file in files:
                        if "fpspr" in file.lower():
                            full_path = os.path.join(root, file)
                            print(f"  ✓ 발견: {full_path}")

                    # 너무 깊이 들어가지 않기 (성능 문제)
                    if root.count(os.sep) - search_path.count(os.sep) > 3:
                        dirs[:] = []

        print("\n=== 5. UI Automation으로 컨트롤 속성 확인 ===")

        try:
            import comtypes
            from comtypes.client import CreateObject

            # UI Automation 초기화
            uia = CreateObject("UIAutomationCore.CUIAutomation", interface=None)

            # HWND로부터 Element 얻기
            element = uia.ElementFromHandle(hwnd)

            if element:
                # 컨트롤 타입 정보
                try:
                    control_type = element.CurrentControlType
                    class_name = element.CurrentClassName
                    automation_id = element.CurrentAutomationId

                    print(f"  ControlType: {control_type}")
                    print(f"  ClassName: {class_name}")
                    print(f"  AutomationId: {automation_id}")

                    # 패턴 확인
                    patterns = [
                        (10000, "Invoke"),
                        (10001, "Selection"),
                        (10002, "Value"),
                        (10003, "RangeValue"),
                        (10004, "Scroll"),
                        (10005, "ExpandCollapse"),
                        (10006, "Grid"),
                        (10007, "GridItem"),
                        (10010, "Text"),
                        (10014, "Table"),
                        (10015, "TableItem"),
                    ]

                    print("\n  지원하는 패턴:")
                    for pattern_id, pattern_name in patterns:
                        try:
                            pattern_ptr = element.GetCurrentPattern(pattern_id)
                            if pattern_ptr:
                                print(f"    ✓ {pattern_name}")
                        except:
                            pass

                except Exception as e:
                    print(f"  UIA 속성 조회 실패: {e}")
        except Exception as e:
            print(f"UI Automation 실패: {e}")

        capture_func("attempt68_01_complete.png")

        # 결과 요약
        summary = f"""
검색 결과:
- Fpspr OCX: {len(ocx_paths)}개
- ProgID: {len(found_progids)}개
- TypeLib: 확인 필요
"""

        return {
            "success": False,
            "message": f"포괄적 검색 완료\n{summary}"
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"오류: {e}\n{traceback.format_exc()}"
        }
