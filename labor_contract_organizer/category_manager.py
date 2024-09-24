import os
import win32com.client as win32

def organize_by_category(contracts, category, excel_instance):
    """
    사원별 계약서를 카테고리별로 분류하여 저장.
    근로계약서와 시급제계약서 폴더를 각각 구분하여 저장.
    """
    for contract_type, contract_list in contracts.items():
        for employee_name, contract_sheet in contract_list:
            try:
                if contract_sheet is None:
                    print(f"Error: {employee_name}의 계약서가 올바르게 로드되지 않았습니다.")
                    continue

                # 카테고리 값 추출
                if contract_type == "근로계약서":
                    if category == "담당 업무":
                        category_value = contract_sheet.Range("AB8").Value or "Unknown"
                    elif category == "근무 장소":
                        category_value = contract_sheet.Range("AB3").Value or "Unknown"
                    elif category == "시급":
                        category_value = contract_sheet.Range("AB16").Value or "Unknown"
                else:  # 시급제계약서
                    if category == "담당 업무":
                        category_value = contract_sheet.Range("AB8").Value or "Unknown"
                    elif category == "근무 장소":
                        category_value = contract_sheet.Range("AB4").Value or "Unknown"
                    elif category == "시급":
                        category_value = contract_sheet.Range("AB11").Value or "Unknown"

                # 근로계약서/시급제계약서 폴더 생성
                folder_path = os.path.join("uploads/output", contract_type,f"{category}", f"{category_value}")
                os.makedirs(folder_path, exist_ok=True)

                # 계약서 저장
                save_excel_as_is(contract_sheet, folder_path, employee_name, excel_instance)

            except Exception as e:
                print(f"Error processing {employee_name}: {e}")
                continue

def save_excel_as_is(contract_sheet, folder_path, employee_name, excel_instance):
    """
    계약서를 복사하여 엑셀 파일로 저장
    """
    try:
        if contract_sheet is None:
            print(f"Error: {employee_name}의 계약서 시트가 비어있습니다.")
            return

        # 절대 경로로 변환
        folder_path = os.path.abspath(folder_path)

        # 파일명 설정
        excel_path = os.path.join(folder_path, f"{employee_name}_contract.xlsx")
        excel_path = os.path.normpath(excel_path)

        # 파일이 이미 존재하면 저장을 건너뛰기
        if os.path.exists(excel_path):
            print(f"파일이 이미 존재합니다: {excel_path}. 저장을 건너뜁니다.")
            return

        # 새로운 워크북 생성 및 시트 복사
        new_workbook = excel_instance.Workbooks.Add()
        contract_sheet.Copy(Before=new_workbook.Sheets(1))

        # 저장 및 워크북 닫기
        new_workbook.SaveAs(excel_path)
        new_workbook.Close(SaveChanges=False)

        print(f"엑셀 파일 저장 완료: {excel_path}")

    except Exception as e:
        print(f"Error saving {employee_name}'s contract: {e}")
