"""
시도 89: pyexcel-xlsx 라이브러리로 엑셀 읽기

pyexcel은 다양한 포맷 지원
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')


def run():
    print("\n" + "="*60)
    print("시도 89: pyexcel-xlsx로 엑셀 읽기")
    print("="*60)

    try:
        import pyexcel as pe
        import pandas as pd
        print("✓ pyexcel 라이브러리 로드 성공")

        print("\n엑셀 파일 읽기 시도...")

        # pyexcel로 읽기
        records = pe.iget_records(file_name='테스트 임직원.xlsx')

        # 리스트로 변환
        data_list = list(records)

        print(f"✓ 파일 읽기 성공!")
        print(f"  행 수: {len(data_list)}")

        if len(data_list) > 0:
            print(f"  열 수: {len(data_list[0])}")

            # DataFrame으로 변환
            df = pd.DataFrame(data_list)

            print(f"\n컬럼 목록 (처음 10개):")
            for i, col in enumerate(df.columns[:10]):
                print(f"  {i+1}. {col}")

            print(f"\n첫 3행:")
            print(df.head(3).to_string())

            return {
                "success": True,
                "message": f"pyexcel로 읽기 성공! {len(data_list)}행",
                "data": df
            }
        else:
            return {
                "success": False,
                "message": "데이터가 비어있습니다."
            }

    except ImportError as e:
        return {
            "success": False,
            "message": f"pyexcel 미설치: {e}\n\n설치: pip install pyexcel pyexcel-xlsx"
        }
    except Exception as e:
        import traceback
        return {
            "success": False,
            "message": f"pyexcel 실패: {e}\n{traceback.format_exc()}"
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
