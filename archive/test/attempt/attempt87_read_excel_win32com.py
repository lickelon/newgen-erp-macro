"""엑셀 파일 구조 확인 - pyexcel 사용"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

try:
    print("=" * 70)
    print("시도: win32com으로 엑셀 읽기")
    print("=" * 70)

    import win32com.client
    import pandas as pd

    # Excel 애플리케이션 시작
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # 파일 열기
    import os
    filepath = os.path.abspath('테스트 임직원.xlsx')
    print(f"\n파일 경로: {filepath}")

    wb = excel.Workbooks.Open(filepath)
    ws = wb.ActiveSheet

    # 사용된 범위 확인
    used_range = ws.UsedRange
    rows = used_range.Rows.Count
    cols = used_range.Columns.Count

    print(f"\n✓ 파일 열기 성공!")
    print(f"  시트 이름: {ws.Name}")
    print(f"  행 수: {rows}")
    print(f"  열 수: {cols}")

    # 데이터 읽기
    data = []
    for i in range(1, min(rows + 1, 11)):  # 처음 10행만
        row_data = []
        for j in range(1, cols + 1):
            cell_value = ws.Cells(i, j).Value
            row_data.append(cell_value)
        data.append(row_data)

    # DataFrame으로 변환
    if len(data) > 1:
        df = pd.DataFrame(data[1:], columns=data[0])

        print(f"\n컬럼 목록:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. {col}")

        print(f"\n첫 5행 데이터:")
        print(df.head().to_string())
    else:
        print("\n⚠️  데이터가 없습니다.")

    # 닫기
    wb.Close(SaveChanges=False)
    excel.Quit()

    print(f"\n✓ 완료!")

except Exception as e:
    print(f"\n❌ 오류 발생: {e}")
    import traceback
    traceback.print_exc()

    try:
        excel.Quit()
    except:
        pass
