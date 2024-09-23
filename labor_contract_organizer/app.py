import tkinter as tk
from tkinter import filedialog, messagebox
from contract_extractor import extract_contracts
from category_manager import organize_by_category
import win32com.client as win32

class LaborContractOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Labor Contract Organizer")

        # 엑셀 파일 선택 버튼
        self.upload_button = tk.Button(root, text="엑셀 파일 선택", command=self.select_files)
        self.upload_button.pack()

        # 파일 리스트 라벨
        self.file_list_label = tk.Label(root, text="선택된 파일이 없습니다.")
        self.file_list_label.pack()

        # 카테고리 선택 드롭다운
        self.category_var = tk.StringVar(root)
        self.category_var.set("카테고리 선택")
        self.category_dropdown = tk.OptionMenu(root, self.category_var, "담당 업무", "근무 장소", "시급")
        self.category_dropdown.pack()

        # 처리 시작 버튼
        self.process_button = tk.Button(root, text="처리 시작", command=self.process_files)
        self.process_button.pack()

        # 선택된 파일 목록 저장
        self.file_paths = []

    def select_files(self):
        # 여러 개의 엑셀 파일 선택 가능
        self.file_paths = filedialog.askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx")])
        if self.file_paths:
            self.file_list_label.config(text="\n".join(self.file_paths))
        else:
            self.file_list_label.config(text="선택된 파일이 없습니다.")

    def process_files(self):
        if not self.file_paths:
            messagebox.showerror("오류", "엑셀 파일을 선택하세요!")
            return
        
        category = self.category_var.get()
        if category == "카테고리 선택":
            messagebox.showerror("오류", "카테고리를 선택하세요!")
            return

        # Excel 인스턴스 생성
        excel_instance = win32.Dispatch("Excel.Application")
        excel_instance.Visible = False
        excel_instance.DisplayAlerts = False
        
        try:
            # 선택된 모든 파일을 처리
            for file_path in self.file_paths:
                print(f"Processing file: {file_path}")
                contracts = extract_contracts(file_path, excel_instance)
                organize_by_category(contracts, category, excel_instance)
        finally:
            excel_instance.Quit()

        messagebox.showinfo("완료", "모든 파일 처리가 완료되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = LaborContractOrganizerApp(root)
    root.mainloop()
