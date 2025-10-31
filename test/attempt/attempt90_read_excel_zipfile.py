"""
시도 90: zipfile + xml.etree로 직접 xlsx 파싱

xlsx는 사실 zip 압축된 XML 파일들의 모음
스타일 정보를 무시하고 데이터만 추출
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')


def run():
    print("\n" + "="*60)
    print("시도 90: zipfile로 xlsx 직접 파싱")
    print("="*60)

    try:
        import zipfile
        import xml.etree.ElementTree as ET
        import pandas as pd
        print("✓ 라이브러리 로드 성공")

        print("\n엑셀 파일 읽기 시도...")

        # xlsx를 zip으로 열기
        with zipfile.ZipFile('테스트 임직원.xlsx', 'r') as zip_ref:
            # 파일 목록 확인
            print(f"  Zip 파일 목록:")
            for name in zip_ref.namelist()[:10]:
                print(f"    - {name}")

            # sharedStrings.xml 읽기 (문자열 테이블)
            try:
                with zip_ref.open('xl/sharedStrings.xml') as f:
                    strings_content = f.read()
                    strings_tree = ET.fromstring(strings_content)

                    # 모든 문자열 추출
                    ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    strings = []
                    for si in strings_tree.findall('.//t', ns):
                        strings.append(si.text if si.text else '')

                    print(f"\n  ✓ SharedStrings: {len(strings)}개")
            except:
                strings = []
                print(f"\n  ✗ SharedStrings 없음")

            # sheet1.xml 읽기
            with zip_ref.open('xl/worksheets/sheet1.xml') as f:
                sheet_content = f.read()
                sheet_tree = ET.fromstring(sheet_content)

                ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

                # 모든 행 추출
                rows_data = []
                for row in sheet_tree.findall('.//row', ns)[:6]:  # 처음 6행만
                    row_data = []
                    for cell in row.findall('.//c', ns):
                        cell_type = cell.get('t')
                        value_elem = cell.find('.//v', ns)

                        if value_elem is not None:
                            value = value_elem.text

                            # 문자열 타입이면 sharedStrings에서 찾기
                            if cell_type == 's' and value and strings:
                                try:
                                    value = strings[int(value)]
                                except:
                                    pass

                            row_data.append(value)
                        else:
                            row_data.append('')

                    rows_data.append(row_data)

                print(f"\n  ✓ 행 수: {len(rows_data)}")
                print(f"  ✓ 첫 행 열 수: {len(rows_data[0]) if rows_data else 0}")

                if len(rows_data) > 1:
                    # DataFrame으로 변환
                    max_cols = max(len(row) for row in rows_data)

                    # 열 수 맞추기
                    for row in rows_data:
                        while len(row) < max_cols:
                            row.append('')

                    df = pd.DataFrame(rows_data[1:], columns=rows_data[0])

                    print(f"\n컬럼 목록 (처음 10개):")
                    for i, col in enumerate(df.columns[:10]):
                        print(f"  {i+1}. {col}")

                    print(f"\n첫 3행:")
                    print(df.head(3).iloc[:, :5].to_string())  # 처음 5열만

                    return {
                        "success": True,
                        "message": f"zipfile 직접 파싱 성공! {len(rows_data)}행",
                        "data": df
                    }
                else:
                    return {
                        "success": False,
                        "message": "데이터가 비어있습니다."
                    }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"zipfile 파싱 실패: {e}\n{traceback.format_exc()}"
        }


if __name__ == "__main__":
    result = run()
    print("\n" + "="*60)
    print(f"결과: {result['message']}")
    print("="*60)

    if result["success"]:
        print("\n✅ 성공!")
    else:
        print("\n⚠️  실패")
