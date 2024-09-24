import win32com.client as win32
import re

def extract_contracts(file_path, excel_instance):
    """
    엑셀 파일에서 근로계약서와 시급제 계약서를 추출하여 리스트로 반환.
    """
    contracts = {
        "근로계약서": [],
        "시급제계약서": []
    }

    # 원본 엑셀 파일 열기
    workbook = excel_instance.Workbooks.Open(file_path)
    
    # 급여설계 시트와 근로계약서, 시급제계약서 시트 로드
    try:
        salary_sheet = workbook.Sheets("급여설계")
        contract_sheet = workbook.Sheets("1. 근로계약서")
        hourly_contract_sheet = workbook.Sheets("1-1. 시급제계약서")
    except Exception as e:
        print(f"Error loading sheets: {e}")
        workbook.Close(SaveChanges=False)
        return contracts

    # 연번을 기준으로 근로계약서와 시급제 계약서를 구분
    row = 5  # A5부터 시작
    while True:
        employee_number = salary_sheet.Cells(row, 1).Text  # A열의 연번
        employee_name = salary_sheet.Cells(row, 3).Value    # C열의 사원 이름
        
        # 연번이 None이거나 더 이상 데이터가 없으면 루프 종료
        if employee_number is None or str(employee_number).strip() == "":
            print(f"end: {row}")
            break
        
        # 연번이 숫자로만 된 경우에는 근로계약서 추출
        if re.match(r"^\d+$", str(employee_number).strip()):  # 숫자인지 확인
            print(f"근로계약서 존재")
            if employee_name:
                # 근로계약서 시트에서 사원의 이름을 AB6 셀에 입력
                contract_sheet.Range("AB6").Value = employee_name

                # 새 워크북에 시트를 복사
                contract_sheet.Copy()  # 시트를 복사
                new_workbook = excel_instance.Workbooks.Item(excel_instance.Workbooks.Count)  # 가장 최근에 생성된 워크북을 참조
                new_contract_sheet = new_workbook.Sheets(1)  # 새로 복사된 시트 가져오기

                # 근로계약서 목록에 추가
                contracts["근로계약서"].append((employee_name, new_contract_sheet))

        # 연번이 시급제계약서인 경우
        elif re.match(r"^시급\d+$", str(employee_number)):  # 시급1, 시급2 등 확인
            print(f"시급제계약서 존재")
            if employee_name:
                # 시급제계약서 시트에서 사원의 이름을 AB6 셀에 입력
                hourly_contract_sheet.Range("AB6").Value = employee_name

                # 새 워크북에 시트를 복사
                hourly_contract_sheet.Copy()  # 시트를 복사
                new_workbook = excel_instance.Workbooks.Item(excel_instance.Workbooks.Count)  # 가장 최근에 생성된 워크북을 참조
                new_contract_sheet = new_workbook.Sheets(1)  # 새로 복사된 시트 가져오기

                # 시급제계약서 목록에 추가
                contracts["시급제계약서"].append((employee_name, new_contract_sheet))
        
        row += 1  # 다음 행으로 이동

    # 원본 엑셀 파일 닫기
    workbook.Close(SaveChanges=False)

    return contracts
