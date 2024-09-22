import os
import time
import win32com.client as win32

def organize_by_category(contracts, category, excel_instance):
    """
    사원별 계약서를 pywin32를 사용하여 카테고리별로 분류하여 저장
    contracts: 사원 이름과 근로계약서 시트
    category: '담당 업무', '근무 장소', '시급' 중 하나
    file_path: 원본 엑셀 파일 경로
    """
    for employee_name, contract_sheet in contracts:
        try:
            # contract_sheet가 None인지 체크
            if contract_sheet is None:
                print(f"Error: {employee_name}의 계약서가 올바르게 로드되지 않았습니다.")
                continue
            
            # 카테고리별로 값을 추출
            if category == "담당 업무":
                category_value = contract_sheet.Range("AB8").Value or "Unknown"
            elif category == "근무 장소":
                category_value = contract_sheet.Range("AB3").Value or "Unknown"
            elif category == "시급":
                category_value = contract_sheet.Range("AB16").Value or "Unknown"
            else:
                continue

            # 카테고리별로 폴더 생성
            folder_path = os.path.join("uploads/output", f"{category}_{category_value}")
            os.makedirs(folder_path, exist_ok=True)

            # 계약서를 해당 폴더에 엑셀로 저장
            save_excel_as_is(contract_sheet, folder_path, employee_name, excel_instance)

        except Exception as e:
            print(f"Error processing {employee_name}: {e}")
            continue


def save_excel_as_is(contract_sheet, folder_path, employee_name, excel_instance):
    """
    pywin32를 사용하여 근로계약서를 그대로 복사하여 새 엑셀 파일로 저장
    """
    try:
        # contract_sheet가 None이 아닌지 체크
        if contract_sheet is None:
            print(f"Error: {employee_name}의 계약서 시트가 비어있습니다.")
            return

        # 절대 경로로 변환
        folder_path = os.path.abspath(folder_path)

        # 파일 저장 경로 설정 (타임스탬프 미사용)
        excel_path = os.path.join(folder_path, f"{employee_name}_contract.xlsx")
        excel_path = os.path.normpath(excel_path)  # 경로 정규화하여 슬래시와 백슬래시 문제 해결

        # 파일이 이미 존재하는지 확인
        if os.path.exists(excel_path):
            print(f"파일이 이미 존재합니다: {excel_path}. 저장을 건너뜁니다.")
            return  # 파일이 존재하면 저장하지 않고 함수 종료

        # 디버깅을 위해 경로 출력
        print(f"Saving Excel to: {excel_path}")

        # 저장 경로가 존재하는지 확인
        os.makedirs(folder_path, exist_ok=True)

        # 새로운 워크북 생성 및 시트 복사 수행
        new_workbook = excel_instance.Workbooks.Add()
        contract_sheet.Copy(Before=new_workbook.Sheets(1))

        # 새로운 워크북 저장 및 닫기
        new_workbook.SaveAs(excel_path)
        new_workbook.Close(SaveChanges=False)

        print(f"엑셀 파일 저장 완료: {excel_path}")

    except Exception as e:
        print(f"Error saving {employee_name}'s contract: {e}")
