import win32com.client as win32
import re  # 정규 표현식을 사용하여 연번에서 숫자만 추출

def extract_contracts(file_path, excel_instance):
    """
    엑셀 파일에서 숫자로 된 연번(근로계약서)을 가진 사원들의 계약서를 추출하여 리스트로 반환
    """
    employee_contracts = []

    # 원본 엑셀 파일 열기
    workbook = excel_instance.Workbooks.Open(file_path)
    
    # 급여설계 시트와 근로계약서 시트 로드
    try:
        salary_sheet = workbook.Sheets("급여설계")
        contract_sheet = workbook.Sheets("1. 근로계약서")
    except Exception as e:
        print(f"Error loading sheets: {e}")
        workbook.Close(SaveChanges=False)
        return []

    # 연번 범위를 동적으로 설정하여 숫자로 된 연번만 추출
    row = 5  # A5부터 시작
    while True:
        employee_number = salary_sheet.Cells(row, 1).Text  # A열의 연번
        employee_name = salary_sheet.Cells(row, 3).Value    # C열의 사원 이름
        
        # 연번이 없는 경우 루프 종료
        if employee_number is None or str(employee_number).strip() == "":
            break
        
        # 연번이 숫자로만 된 경우에만 계약서 추출 (근로계약서인 경우)
        if re.match(r"^\d+$", str(employee_number).strip()):  # 숫자인지 확인
            if employee_name:
                # 근로계약서 시트에서 사원의 이름을 AB6 셀에 입력
                contract_sheet.Range("AB6").Value = employee_name

                # 새 워크북에 시트를 복사
                contract_sheet.Copy()  # 시트를 복사
                new_workbook = excel_instance.Workbooks.Item(excel_instance.Workbooks.Count)  # 가장 최근에 생성된 워크북을 참조
                new_contract_sheet = new_workbook.Sheets(1)  # 새로 복사된 시트 가져오기

                # 사원의 이름과 복사된 계약서 시트를 저장
                employee_contracts.append((employee_name, new_contract_sheet))
        
        row += 1  # 다음 행으로 이동

    # 원본 엑셀 파일 닫기
    workbook.Close(SaveChanges=False)

    return employee_contracts
