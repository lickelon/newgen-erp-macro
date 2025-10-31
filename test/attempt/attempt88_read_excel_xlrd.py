"""
시도 88: xlrd 라이브러리로 엑셀 읽기

xlrd는 .xls 파일용이지만 일부 xlsx도 지원
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')


def run():
    print("\n" + "="*60)
    print("시도 88: xlrd로 엑셀 읽기")
    print("="*60)

    try:
        import xlrd
        import pandas as pd
        print("✓ xlrd 라이브러리 로드 성공")

        print("\n엑셀 파일 읽기 시도...")

        # xlrd로 직접 읽기
        wb = xlrd.open_workbook('테스트 임직원.xlsx')
        ws = wb.sheet_by_index(0)

        print(f"✓ 파일 열기 성공!")
        print(f"  시트 이름: {ws.name}")
        print(f"  행 수: {ws.nrows}")
        print(f"  열 수: {ws.ncols}")

        # 데이터 읽기
        data = []
        for row_idx in range(min(ws.nrows, 6)):  # 처음 6행
            row_data = []
            for col_idx in range(ws.ncols):
                cell = ws.cell(row_idx, col_idx)
                row_data.append(cell.value)
            data.append(row_data)

        # DataFrame으로 변환
        df = pd.DataFrame(data[1:], columns=data[0])

        print(f"\n컬럼 목록 (처음 10개):")
        for i, col in enumerate(df.columns[:10]):
            print(f"  {i+1}. {col}")

        print(f"\n첫 3행:")
        print(df.head(3).to_string())

        return {
            "success": True,
            "message": f"xlrd로 읽기 성공! {ws.nrows}행 × {ws.ncols}열",
            "data": df
        }

    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"xlrd 실패: {e}\n{traceback.format_exc()}"
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
