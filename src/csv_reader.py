"""
CSV 테스트 데이터 읽기 모듈
부양가족 정보 CSV 파일을 읽고 파싱합니다.
"""

import csv
from typing import List, Dict
from collections import defaultdict


class DependentData:
    """부양가족 한 명의 데이터"""

    def __init__(self, row: Dict[str, str]):
        self.employee_no = row['근로자\n사번'].strip()
        self.employee_name = row['근로자명'].strip()
        self.relationship_code = row['관계코드'].strip()
        self.name = row['이름'].strip()
        self.nationality = row['내/외국인'].strip()
        self.id_number = row['주민등록번호'].strip()
        self.age = row['만나이'].strip()
        self.basic_deduction = row['기본공제여부'].strip()
        self.child_deduction = row['자녀공제'].strip()

    def __repr__(self):
        return f"DependentData({self.employee_no}, {self.name}, 관계={self.relationship_code})"

    def to_dict(self) -> Dict[str, str]:
        """딕셔너리로 변환"""
        return {
            'employee_no': self.employee_no,
            'employee_name': self.employee_name,
            'relationship_code': self.relationship_code,
            'name': self.name,
            'nationality': self.nationality,
            'id_number': self.id_number,
            'age': self.age,
            'basic_deduction': self.basic_deduction,
            'child_deduction': self.child_deduction
        }


class CSVReader:
    """CSV 파일 읽기"""

    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.data: List[DependentData] = []

    def read(self) -> List[DependentData]:
        """CSV 파일을 읽어서 DependentData 리스트로 반환"""
        with open(self.csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                self.data.append(DependentData(row))

        return self.data

    def get_by_employee(self, employee_no: str) -> List[DependentData]:
        """특정 사원의 부양가족 데이터만 반환"""
        return [d for d in self.data if d.employee_no == employee_no]

    def group_by_employee(self) -> Dict[str, List[DependentData]]:
        """사원별로 그룹화된 딕셔너리 반환"""
        grouped = defaultdict(list)
        for d in self.data:
            grouped[d.employee_no].append(d)
        return dict(grouped)

    def get_employee_list(self) -> List[tuple]:
        """사원 목록 반환 [(사번, 이름)]"""
        employees = {}
        for d in self.data:
            if d.relationship_code == '0':  # 본인만
                employees[d.employee_no] = d.employee_name

        return [(no, name) for no, name in employees.items()]

    def print_summary(self):
        """데이터 요약 출력"""
        grouped = self.group_by_employee()
        print(f"총 사원 수: {len(grouped)}")
        print(f"총 부양가족 레코드 수: {len(self.data)}")
        print("\n사원별 부양가족 수:")
        for emp_no, dependents in grouped.items():
            emp_name = dependents[0].employee_name
            print(f"  {emp_no} ({emp_name}): {len(dependents)}명")


if __name__ == "__main__":
    # 테스트
    reader = CSVReader("테스트 데이터.csv")
    data = reader.read()

    print("=== CSV 데이터 읽기 테스트 ===\n")
    reader.print_summary()

    print("\n=== 첫 번째 사원 데이터 ===")
    employees = reader.get_employee_list()
    if employees:
        first_emp_no, first_emp_name = employees[0]
        print(f"\n사원: {first_emp_no} - {first_emp_name}")
        dependents = reader.get_by_employee(first_emp_no)
        for dep in dependents:
            print(f"  {dep}")
