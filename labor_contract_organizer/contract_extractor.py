import win32com.client as win32

def extract_contracts(file_path, excel_instance):
    """
    엑셀 파일에서 각 사원의 근로계약서를 추출하여 리스트로 반환
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

    # A5부터 A9까지 범위 순회하여 사원 이름 추출 (C열에서 사원 이름 가져오기)
    for i in range(5, 10):
        employee_name = salary_sheet.Cells(i, 3).Value
        
        if employee_name:
            # 근로계약서 시트에서 사원의 이름을 AB6 셀에 입력
            contract_sheet.Range("AB6").Value = employee_name

            # 새 워크북에 시트를 복사
            contract_sheet.Copy()  # 시트를 복사
            new_workbook = excel_instance.Workbooks.Item(excel_instance.Workbooks.Count)  # 가장 최근에 생성된 워크북을 참조
            new_contract_sheet = new_workbook.Sheets(1)  # 새로 복사된 시트 가져오기

            # 사원의 이름과 복사된 계약서 시트를 저장
            employee_contracts.append((employee_name, new_contract_sheet))

    # 원본 엑셀 파일 닫기
    workbook.Close(SaveChanges=False)

    return employee_contracts
