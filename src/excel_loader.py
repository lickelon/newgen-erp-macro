"""
엑셀 파일 로더

zipfile을 사용하여 xlsx 파일을 직접 파싱
openpyxl 호환성 문제를 우회
"""
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path


class ExcelLoader:
    """xlsx 파일을 zipfile로 직접 파싱하는 로더"""

    def __init__(self, filepath):
        """
        Args:
            filepath: 엑셀 파일 경로
        """
        self.filepath = Path(filepath)
        if not self.filepath.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    def load(self, sheet_name=None):
        """
        엑셀 파일을 DataFrame으로 로드

        Args:
            sheet_name: 시트 이름 (None이면 첫 번째 시트)

        Returns:
            pandas.DataFrame
        """
        with zipfile.ZipFile(self.filepath, 'r') as zip_ref:
            # SharedStrings 로드
            shared_strings = self._load_shared_strings(zip_ref)

            # Sheet 데이터 로드
            sheet_xml = 'xl/worksheets/sheet1.xml'  # 기본: 첫 번째 시트
            rows_data = self._load_sheet_data(zip_ref, sheet_xml, shared_strings)

            # DataFrame으로 변환
            if len(rows_data) > 1:
                # 최대 열 수 계산
                max_cols = max(len(row) for row in rows_data)

                # 모든 행의 열 수를 맞춤
                for row in rows_data:
                    while len(row) < max_cols:
                        row.append('')

                # 첫 행을 컬럼으로, 나머지를 데이터로
                df = pd.DataFrame(rows_data[1:], columns=rows_data[0])
                return df
            else:
                return pd.DataFrame()

    def _load_shared_strings(self, zip_ref):
        """
        sharedStrings.xml 로드

        Args:
            zip_ref: ZipFile 객체

        Returns:
            list: 문자열 리스트
        """
        try:
            with zip_ref.open('xl/sharedStrings.xml') as f:
                content = f.read()
                tree = ET.fromstring(content)

                ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                strings = []

                for si in tree.findall('.//t', ns):
                    strings.append(si.text if si.text else '')

                return strings
        except KeyError:
            # sharedStrings.xml이 없는 경우
            return []

    def _load_sheet_data(self, zip_ref, sheet_xml, shared_strings):
        """
        시트 데이터 로드

        Args:
            zip_ref: ZipFile 객체
            sheet_xml: 시트 XML 경로
            shared_strings: 문자열 테이블

        Returns:
            list: 행 데이터 리스트
        """
        with zip_ref.open(sheet_xml) as f:
            content = f.read()
            tree = ET.fromstring(content)

            ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

            rows_data = []
            for row in tree.findall('.//row', ns):
                row_data = []

                for cell in row.findall('.//c', ns):
                    cell_type = cell.get('t')
                    value_elem = cell.find('.//v', ns)

                    if value_elem is not None:
                        value = value_elem.text

                        # 문자열 타입이면 sharedStrings에서 찾기
                        if cell_type == 's' and value and shared_strings:
                            try:
                                value = shared_strings[int(value)]
                            except (IndexError, ValueError):
                                pass

                        row_data.append(value)
                    else:
                        row_data.append('')

                rows_data.append(row_data)

            return rows_data


def load_employees(filepath):
    """
    사원 정보 엑셀 로드

    Args:
        filepath: 엑셀 파일 경로

    Returns:
        pandas.DataFrame: 사원 정보

    Raises:
        FileNotFoundError: 파일이 없을 때
        ValueError: 필수 컬럼이 없을 때
    """
    loader = ExcelLoader(filepath)
    df = loader.load()

    # 필수 컬럼 검증
    required_cols = ['사원번호\n(10자리)']  # 실제 컬럼명에 개행 포함
    if not all(col in df.columns for col in required_cols):
        raise ValueError(f"필수 컬럼 누락: {required_cols}")

    # 컬럼명 정리 (개행 제거)
    df.columns = [col.replace('\n', ' ').strip() for col in df.columns]

    return df


def load_dependents(filepath):
    """
    부양가족 정보 엑셀 로드

    Args:
        filepath: 엑셀 파일 경로

    Returns:
        pandas.DataFrame: 부양가족 정보

    Raises:
        FileNotFoundError: 파일이 없을 때
        ValueError: 필수 컬럼이 없을 때
    """
    loader = ExcelLoader(filepath)
    df = loader.load()

    # 컬럼명 정리
    df.columns = [col.replace('\n', ' ').strip() for col in df.columns]

    # 필수 컬럼 검증 (정리된 컬럼명으로)
    required_cols = ['사원번호 (10자리)', '이름', '관계', '주민등록번호']
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        print(f"⚠️  경고: 필수 컬럼 일부 누락: {missing_cols}")
        print(f"사용 가능한 컬럼: {df.columns.tolist()}")

    return df


def get_dependents_for_employee(dependents_df, empno):
    """
    특정 사원의 부양가족 조회

    Args:
        dependents_df: 부양가족 DataFrame
        empno: 사원번호

    Returns:
        pandas.DataFrame: 해당 사원의 부양가족
    """
    empno_col = '사원번호 (10자리)'

    if empno_col not in dependents_df.columns:
        raise ValueError(f"컬럼 '{empno_col}'을(를) 찾을 수 없습니다.")

    return dependents_df[dependents_df[empno_col] == empno]


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    print("=" * 70)
    print("엑셀 로더 테스트")
    print("=" * 70)

    try:
        # 사원 정보 로드
        print("\n1. 사원 정보 로드...")
        employees = load_employees('테스트 임직원.xlsx')

        print(f"✓ 사원 수: {len(employees)}")
        print(f"✓ 컬럼 수: {len(employees.columns)}")

        print(f"\n주요 컬럼:")
        key_cols = ['사원번호 (10자리)', '사원명 (30자리)', '주민(외국인)등록번호']
        for col in key_cols:
            if col in employees.columns:
                print(f"  - {col}")

        print(f"\n첫 5명:")
        if '사원번호 (10자리)' in employees.columns:
            print(employees[['사원번호 (10자리)', '사원명 (30자리)']].head())

        print("\n✅ 테스트 성공!")

    except Exception as e:
        print(f"\n❌ 오류: {e}")
        import traceback
        traceback.print_exc()
