import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import json
from logger import logger
from api_utils import APIHandler
from excel_processor import ExcelHandler

class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Data Processor")
        self.geometry("600x800")
        ctk.set_appearance_mode("light")

        self.source_file_path = ""
        self.form_file_path = ""
        self.signature_dir = ""
        self.naver_client_id = ""
        self.naver_client_secret = ""
        self.ncp_client_id = ""
        self.ncp_client_secret = ""
        
        self.load_settings()
        logger.info("Application started")

        # Main Container
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=40, pady=20)

        # UI Sections
        self.source_label = self.create_section(self.main_frame, "📂 추출을 원하는 엑셀 파일", "파일을 선택하세요", self.select_source_file)
        if self.source_file_path:
            self.source_label.configure(text=os.path.basename(self.source_file_path), text_color="black")
            
        self.add_spacing(self.main_frame, 20)
        self.form_label = self.create_section(self.main_frame, "📋 폼 엑셀 파일", "폼 파일을 선택하세요", self.select_form_file)
        if self.form_file_path:
            self.form_label.configure(text=os.path.basename(self.form_file_path), text_color="black")
            
        self.add_spacing(self.main_frame, 20)
        self.sig_label = self.create_section(self.main_frame, "✍️ 서명 이미지 폴더", "서명 폴더를 선택하세요", self.select_signature_dir)
        if self.signature_dir:
            self.sig_label.configure(text=os.path.basename(self.signature_dir), text_color="black")
            
        self.add_spacing(self.main_frame, 20)
        
        # API Settings Section
        self.api_frame = ctk.CTkFrame(self.main_frame, fg_color="white", corner_radius=0)
        self.api_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(self.api_frame, text="🔑 네이버 API 설정", font=("Malgun Gothic", 16, "bold"), anchor="w").pack(fill="x", padx=10, pady=(10, 5))
        self.api_status_label = ctk.CTkLabel(self.api_frame, text="API 키를 설정해주세요", font=("Malgun Gothic", 12), text_color="#666666")
        self.api_status_label.pack(side="left", padx=10, pady=5)
        ctk.CTkButton(self.api_frame, text="설정", width=60, height=30, fg_color="#1FA1FF", command=self.open_api_settings).pack(side="right", padx=10, pady=5)

        self.add_spacing(self.main_frame, 40)

        # Run Button
        self.run_button = ctk.CTkButton(self.main_frame, text="📂 모든 정보를 설정해주세요", height=45, state="disabled", fg_color="#A0A0A0", text_color="white", command=self.process_excel)
        self.run_button.pack(fill="x")

        self.update_api_status()

    def create_section(self, parent, title_text, placeholder_text, button_command):
        section_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=0)
        section_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(section_frame, text=title_text, font=("Malgun Gothic", 16, "bold"), anchor="w").pack(fill="x", padx=10, pady=(10, 5))
        path_frame = ctk.CTkFrame(section_frame, fg_color="#F0F0F0", height=40, corner_radius=0)
        path_frame.pack(fill="x", padx=10, pady=5)
        path_frame.pack_propagate(False)
        path_label = ctk.CTkLabel(path_frame, text=placeholder_text, font=("Malgun Gothic", 12), text_color="#666666")
        path_label.pack(expand=True)
        ctk.CTkButton(section_frame, text="파일 선택", fg_color="#1FA1FF", hover_color="#0080FF", corner_radius=2, text_color="white", command=lambda: button_command(path_label)).pack(pady=(5, 15))
        return path_label

    def add_spacing(self, parent, size):
        ctk.CTkFrame(parent, height=size, fg_color="transparent").pack()

    def select_source_file(self, label):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.source_file_path = path
            label.configure(text=os.path.basename(path), text_color="black")
            logger.info(f"Source file selected: {path}")
            self.save_settings()
            self.check_files_selected()

    def select_form_file(self, label):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.form_file_path = path
            label.configure(text=os.path.basename(path), text_color="black")
            logger.info(f"Form file selected: {path}")
            self.save_settings()
            self.check_files_selected()

    def select_signature_dir(self, label):
        path = filedialog.askdirectory()
        if path:
            self.signature_dir = path
            label.configure(text=os.path.basename(path), text_color="black")
            logger.info(f"Signature directory selected: {path}")
            self.save_settings()
            self.check_files_selected()

    def open_api_settings(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("네이버 API 설정")
        dialog.geometry("450x450")
        dialog.transient(self)
        
        # 1. Search API (Client ID/Secret)
        ctk.CTkLabel(dialog, text="[검색 API] Naver Client ID:").pack(pady=(20, 5))
        search_id_entry = ctk.CTkEntry(dialog, width=350)
        search_id_entry.insert(0, self.naver_client_id)
        search_id_entry.pack(pady=5)
        ctk.CTkLabel(dialog, text="[검색 API] Naver Client Secret:").pack(pady=(5, 5))
        search_secret_entry = ctk.CTkEntry(dialog, width=350, show="*")
        search_secret_entry.insert(0, self.naver_client_secret)
        search_secret_entry.pack(pady=5)

        # 2. Maps API (NCP ID/Key)
        ctk.CTkLabel(dialog, text="[지도 API] NCP Client ID:").pack(pady=(15, 5))
        ncp_id_entry = ctk.CTkEntry(dialog, width=350)
        ncp_id_entry.insert(0, self.ncp_client_id)
        ncp_id_entry.pack(pady=5)
        ctk.CTkLabel(dialog, text="[지도 API] NCP Client Secret:").pack(pady=(5, 5))
        ncp_secret_entry = ctk.CTkEntry(dialog, width=350, show="*")
        ncp_secret_entry.insert(0, self.ncp_client_secret)
        ncp_secret_entry.pack(pady=5)

        def save():
            self.naver_client_id = search_id_entry.get().strip()
            self.naver_client_secret = search_secret_entry.get().strip()
            self.ncp_client_id = ncp_id_entry.get().strip()
            self.ncp_client_secret = ncp_secret_entry.get().strip()
            self.save_settings()
            self.update_api_status()
            logger.info("API settings updated by user")
            dialog.destroy()
        ctk.CTkButton(dialog, text="저장", command=save).pack(pady=20)

    def update_api_status(self):
        if self.naver_client_id and self.naver_client_secret:
            status = "API 키 설정됨 ✅"
            if self.ncp_client_id and self.ncp_client_secret:
                status = "모든 API 키 설정됨 ✅"
            self.api_status_label.configure(text=status, text_color="green")
        else:
            self.api_status_label.configure(text="API 키를 설정해주세요", text_color="#666666")
        self.check_files_selected()

    def load_settings(self):
        if os.path.exists("settings.json"):
            try:
                with open("settings.json", "r") as f:
                    settings = json.load(f)
                    self.naver_client_id = settings.get("naver_client_id", "")
                    self.naver_client_secret = settings.get("naver_client_secret", "")
                    self.ncp_client_id = settings.get("ncp_client_id", "")
                    self.ncp_client_secret = settings.get("ncp_client_secret", "")
                    self.source_file_path = settings.get("source_file_path", "")
                    self.form_file_path = settings.get("form_file_path", "")
                    self.signature_dir = settings.get("signature_dir", "")
            except Exception as e:
                logger.error(f"Failed to load settings: {e}")

    def save_settings(self):
        try:
            settings = {
                "naver_client_id": self.naver_client_id, 
                "naver_client_secret": self.naver_client_secret,
                "ncp_client_id": self.ncp_client_id,
                "ncp_client_secret": self.ncp_client_secret,
                "source_file_path": self.source_file_path,
                "form_file_path": self.form_file_path,
                "signature_dir": self.signature_dir
            }
            with open("settings.json", "w") as f:
                json.dump(settings, f)
        except Exception as e:
            logger.error(f"Failed to save settings: {e}")

    def check_files_selected(self):
        if not hasattr(self, 'run_button'):
            return
            
        if all([self.source_file_path, self.form_file_path, self.signature_dir, self.naver_client_id]):
            self.run_button.configure(state="normal", text="📊 추출 및 실행하기", fg_color="#1FA1FF")
        else:
            self.run_button.configure(state="disabled", text="📂 모든 정보를 설정해주세요", fg_color="#A0A0A0")

    def process_excel(self):
        try:
            logger.info("Button 'Execute' clicked. Starting process...")
            api_handler = APIHandler(
                self.naver_client_id, 
                self.naver_client_secret,
                ncp_client_id=self.ncp_client_id,
                ncp_client_secret=self.ncp_client_secret
            )
            excel_handler = ExcelHandler(self.source_file_path, self.form_file_path, self.signature_dir, api_handler)
            
            count, folder = excel_handler.process()
            
            messagebox.showinfo("완료", f"데이터 처리가 완료되었습니다!\n총 {count}개의 파일이 '{folder}' 폴더에 저장되었습니다.")
            logger.info("Excel processing successful.")
        except Exception as e:
            logger.critical(f"Process failed: {e}", exc_info=True)
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다. 로그 파일을 확인해주세요.\n\n{str(e)}")

if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()
