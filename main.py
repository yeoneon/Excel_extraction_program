import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import os

class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Data Processor")
        self.geometry("600x500")
        ctk.set_appearance_mode("light")  # Matching the image's light theme

        self.source_file_path = ""
        self.form_file_path = ""

        # Main Container
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=40, pady=20)

        # Section 1: 추출을 원하는 엑셀 파일 (Source B)
        self.create_section(
            self.main_frame, 
            "📂 추출을 원하는 엑셀 파일", 
            "파일을 선택하세요", 
            self.select_source_file
        )

        # Spacing
        self.add_spacing(self.main_frame, 20)

        # Section 2: 폼 엑셀 파일 (Form A)
        self.create_section(
            self.main_frame, 
            "📋 폼 엑셀 파일", 
            "폼 파일을 선택하세요", 
            self.select_form_file
        )

        # Spacing
        self.add_spacing(self.main_frame, 40)

        # Section 3: 실행 버튼
        self.run_button = ctk.CTkButton(
            self.main_frame, 
            text="📂 파일을 선택해주세요", 
            height=45, 
            state="disabled",
            fg_color="#A0A0A0",  # Gray when disabled
            text_color="white",
            command=self.process_excel
        )
        self.run_button.pack(fill="x")

    def create_section(self, parent, title_text, placeholder_text, button_command):
        # Container for the section
        section_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=0)
        section_frame.pack(fill="x", pady=5)

        # Title
        title_label = ctk.CTkLabel(
            section_frame, 
            text=title_text, 
            font=("Malgun Gothic", 16, "bold"), 
            anchor="w"
        )
        title_label.pack(fill="x", padx=10, pady=(10, 5))

        # Path Display (Gray box)
        path_frame = ctk.CTkFrame(section_frame, fg_color="#F0F0F0", height=40, corner_radius=0)
        path_frame.pack(fill="x", padx=10, pady=5)
        path_frame.pack_propagate(False)

        path_label = ctk.CTkLabel(
            path_frame, 
            text=placeholder_text, 
            font=("Malgun Gothic", 12), 
            text_color="#666666"
        )
        path_label.pack(expand=True)

        # Select Button
        select_button = ctk.CTkButton(
            section_frame, 
            text="파일 선택", 
            fg_color="#1FA1FF", 
            hover_color="#0080FF",
            corner_radius=2,
            text_color="white",
            command=lambda: button_command(path_label)
        )
        select_button.pack(pady=(5, 15))

        return path_label

    def add_spacing(self, parent, size):
        spacer = ctk.CTkFrame(parent, height=size, fg_color="transparent")
        spacer.pack()

    def select_source_file(self, label):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.source_file_path = file_path
            label.configure(text=os.path.basename(file_path), text_color="black")
            self.check_files_selected()

    def select_form_file(self, label):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.form_file_path = file_path
            label.configure(text=os.path.basename(file_path), text_color="black")
            self.check_files_selected()

    def check_files_selected(self):
        if self.source_file_path and self.form_file_path:
            self.run_button.configure(
                state="normal", 
                text="📊 추출 및 실행하기", 
                fg_color="#1FA1FF"
            )

    def process_excel(self):
        try:
            # 1. Read Source B (using pandas)
            # df = pd.read_excel(self.source_file_path)
            
            # 2. Load Form A (using openpyxl to keep styles)
            # wb = openpyxl.load_workbook(self.form_file_path)
            # ws = wb.active
            
            # TODO: Add logic to extract from df and fill into ws
            
            # 3. Save as C
            # save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            # if save_path:
            #     wb.save(save_path)
            #     messagebox.showinfo("완료", "파일이 성공적으로 생성되었습니다!")
            
            messagebox.showinfo("작업 시작", "데이터 처리를 시작합니다. (추후 로직 연결 예정)")
            
        except Exception as e:
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()
