"""
Gmail_GUI.py 대량 Gmail 보내기 Ver 1.1_250614
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import os
import pandas as pd
import threading
import time


class GmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail 발송 프로그램 v1.5")
        self.root.geometry("1050x750")  # 창 크기 변경
        self.root.resizable(True, True)
        self.recipients_data = []
        self.attachment_files = []
        self.total_attachment_size = 0
        self.is_sending = False
        self.create_widgets()
        self.center_window()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        main_frame = ttk.Frame(scrollable_frame, padding="15")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # --- 2단 레이아웃을 위한 Grid 설정 ---
        main_frame.columnconfigure(0, weight=3)  # 왼쪽 열 가중치
        main_frame.columnconfigure(1, weight=2)  # 오른쪽 열 가중치
        main_frame.rowconfigure(1, weight=1)

        # --- 제목 ---
        title_label = ttk.Label(main_frame, text="📧 Gmail 대량 발송 프로그램", font=('맑은 고딕', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))

        # --- 왼쪽 열 컨테이너 프레임 ---
        left_column_frame = ttk.Frame(main_frame)
        left_column_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_column_frame.rowconfigure(2, weight=1)  # 메일 내용 프레임이 세로로 확장되도록 설정
        left_column_frame.columnconfigure(0, weight=1)

        # --- 오른쪽 열 컨테이너 프레임 ---
        right_column_frame = ttk.Frame(main_frame)
        right_column_frame.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        right_column_frame.rowconfigure(0, weight=1)  # 첨부파일 프레임이 세로로 확장되도록 설정
        right_column_frame.columnconfigure(0, weight=1)

        # --- 왼쪽 열 위젯들 ---

        # 발신자 정보
        sender_frame = ttk.LabelFrame(left_column_frame, text="📤 발신자 정보", padding="15")
        sender_frame.grid(row=0, column=0, sticky="ew", pady=2)
        sender_frame.columnconfigure(1, weight=1)
        ttk.Label(sender_frame, text="Gmail 주소:", font=('맑은 고딕', 11, 'bold')).grid(row=0, column=0, sticky=tk.W, padx=5)
        self.sender_email = tk.StringVar(value="sample@gmail.com")
        email_entry = ttk.Entry(sender_frame, textvariable=self.sender_email, width=15, font=('맑은 고딕', 11), state="readonly")
        email_entry.grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Label(sender_frame, text="앱 비밀번호:", font=('맑은 고딕', 11, 'bold')).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.sender_password = tk.StringVar(value="aaaa bbbb cccc dddd")  # 앱 비밀번호 입력
        self.password_entry = ttk.Entry(sender_frame, textvariable=self.sender_password, show="*", width=15, font=('맑은 고딕', 11), state="readonly")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.show_password = tk.BooleanVar()
        show_btn = ttk.Checkbutton(sender_frame, text="비밀번호 표시", variable=self.show_password, command=self.toggle_password)
        show_btn.grid(row=1, column=2, padx=5)

        # 수신자 정보
        recipient_frame = ttk.LabelFrame(left_column_frame, text="📋 수신자 정보", padding="10")
        recipient_frame.grid(row=1, column=0, sticky="ew", pady=2)
        recipient_frame.columnconfigure(0, weight=1)
        file_btn_frame = ttk.Frame(recipient_frame)
        file_btn_frame.grid(row=0, column=0, sticky="ew", pady=5)
        ttk.Button(file_btn_frame, text="📁 Excel 파일 선택", command=self.load_recipients).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_btn_frame, text="📊 샘플 파일 생성", command=self.create_sample_excel).pack(side=tk.LEFT, padx=5)
        self.excel_file_label = ttk.Label(recipient_frame, text="📄 선택된 파일: 없음", font=('맑은 고딕', 11))
        self.excel_file_label.grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        tree_frame = ttk.Frame(recipient_frame)
        tree_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        tree_frame.columnconfigure(0, weight=1)
        self.recipients_tree = ttk.Treeview(tree_frame, columns=('name', 'email'), show='headings', height=6)
        self.recipients_tree.heading('name', text='👤 이름')
        self.recipients_tree.heading('email', text='📧 이메일')
        self.recipients_tree.column('name', width=200)
        self.recipients_tree.column('email', width=300)
        recipients_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.recipients_tree.yview)
        self.recipients_tree.configure(yscrollcommand=recipients_scrollbar.set)
        self.recipients_tree.grid(row=0, column=0, sticky="nsew")
        recipients_scrollbar.grid(row=0, column=1, sticky="ns")

        # 메일 내용
        content_frame = ttk.LabelFrame(left_column_frame, text="✉️ 메일 내용", padding="15")
        content_frame.grid(row=2, column=0, sticky="nsew", pady=2)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(1, weight=1)  # 본문 ScrolledText가 확장되도록 설정
        ttk.Label(content_frame, text="제목:", font=('맑은 고딕', 11, 'bold')).grid(row=0, column=0, sticky=tk.W, padx=5)
        self.subject = tk.StringVar(value="파이썬으로 Gmail 보내기")
        subject_entry = ttk.Entry(content_frame, textvariable=self.subject, font=('맑은 고딕', 11))
        subject_entry.grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Label(content_frame, text="본문:", font=('맑은 고딕', 11, 'bold')).grid(row=1, column=0, sticky=(tk.W, tk.N), padx=5, pady=2)
        self.body_text = scrolledtext.ScrolledText(content_frame, height=10, font=('맑은 고딕', 11), wrap=tk.WORD)
        self.body_text.grid(row=1, column=1, padx=5, pady=2, sticky="nsew")
        self.body_text.insert(tk.END, "안녕하세요, {name}님!\n\n이 메일은 파이썬으로 작성되었습니다.\n\n감사합니다.")

        # 첨부파일
        attachment_frame = ttk.LabelFrame(right_column_frame, text="📎 첨부파일", padding="10")
        attachment_frame.grid(row=0, column=0, sticky="nsew", pady=2)
        attachment_frame.columnconfigure(0, weight=1)
        attachment_frame.rowconfigure(1, weight=1)  # Listbox가 확장되도록 설정
        attach_btn_frame = ttk.Frame(attachment_frame)
        attach_btn_frame.grid(row=0, column=0, sticky="ew", columnspan=2)
        ttk.Button(attach_btn_frame, text="📎 파일 추가", command=self.add_attachments).pack(side=tk.LEFT, padx=5)
        ttk.Button(attach_btn_frame, text="🗑️ 파일 제거", command=self.remove_attachment).pack(side=tk.LEFT, padx=5)
        self.attachment_listbox = tk.Listbox(attachment_frame, height=5, font=('맑은 고딕', 11))
        self.attachment_listbox.grid(row=1, column=0, pady=10, sticky="nsew", columnspan=2)
        self.total_size_label = ttk.Label(attachment_frame, text="총 용량: 0.0 MB / 25.0 MB", font=('맑은 고딕', 11))
        self.total_size_label.grid(row=2, column=0, sticky="w", padx=5, columnspan=2)

        # 발송 제어
        send_frame = ttk.LabelFrame(right_column_frame, text="🚀 발송 제어", padding="10")
        send_frame.grid(row=1, column=0, sticky="ew", pady=2)
        send_frame.columnconfigure(0, weight=1)
        status_frame = ttk.Frame(send_frame)
        status_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew", columnspan=2)
        self.progress_var = tk.StringVar(value="📋 대기 중...")
        self.progress_label = ttk.Label(status_frame, textvariable=self.progress_var, font=('맑은 고딕', 11, 'bold'))
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))
        self.progress_bar = ttk.Progressbar(status_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, expand=True)
        button_container = ttk.Frame(send_frame)
        button_container.grid(row=1, column=0, padx=5, pady=10, sticky="ew", columnspan=2)
        button_container.columnconfigure(0, weight=1)
        self.send_button = ttk.Button(button_container, text="📧 이메일 보내기", command=self.send_email_start)
        self.send_button.grid(row=0, column=0, columnspan=2, pady=2, ipadx=20, ipady=10, sticky="ew")
        ttk.Button(button_container, text="❌ 초기화", command=self.reset_form).grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(button_container, text="⏹️ 발송 중지", command=self.stop_sending).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # --- 전체 스크롤 설정 ---
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def toggle_password(self):
        show_char = "" if self.show_password.get() else "*"
        self.password_entry.config(show=show_char)

    def create_sample_excel(self):
        try:
            sample_data = {'이름': ['윤용수', '용수윤'],
                           '이메일': ['samths@naver.com', 'samths@kakao.com']}
            df = pd.DataFrame(sample_data)
            file_path = filedialog.asksaveasfilename(title="샘플 파일 저장", defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("성공", f"샘플 파일이 생성되었습니다:\n{file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"샘플 파일 생성 중 오류: {str(e)}")

    def load_recipients(self):
        file_path = filedialog.askopenfilename(title="수신자 Excel 파일 선택",
                                               filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
            required_columns = ['이름', '이메일']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("오류", "Excel 파일에 '이름', '이메일' 컬럼이 필요합니다.")
                return
            self.recipients_data.clear()
            self.recipients_tree.delete(*self.recipients_tree.get_children())
            valid_count = 0
            for _, row in df.iterrows():
                if pd.notna(row['이름']) and pd.notna(row['이메일']):
                    name, email = str(row['이름']).strip(), str(row['이메일']).strip()
                    if name and '@' in email and '.' in email.split('@')[1]:
                        self.recipients_data.append({'name': name, 'email': email})
                        self.recipients_tree.insert('', 'end', values=(name, email))
                        valid_count += 1
            self.excel_file_label.config(text=f"📄 선택된 파일: {os.path.basename(file_path)} ({valid_count}명)")
            messagebox.showinfo("성공", f"{valid_count}명의 유효한 수신자를 로드했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 로드 중 오류 발생: {str(e)}")

    def add_attachments(self):
        files = filedialog.askopenfilenames(title="첨부파일 선택", filetypes=[("All files", "*.*")])
        if not files:
            return
        for file_path in files:
            if file_path in self.attachment_files:
                continue
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if self.total_attachment_size + file_size_mb > 25:
                messagebox.showwarning("용량 초과", "총 첨부파일 용량이 25MB를 초과할 수 없습니다.")
                break
            self.total_attachment_size += file_size_mb
            self.attachment_files.append(file_path)
            self.attachment_listbox.insert(tk.END, f"{os.path.basename(file_path)} ({file_size_mb:.2f}MB)")
        self.update_total_size_label()

    def remove_attachment(self):
        selection = self.attachment_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        file_path = self.attachment_files.pop(index)
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        self.total_attachment_size -= file_size_mb
        self.attachment_listbox.delete(index)
        self.update_total_size_label()

    def update_total_size_label(self):
        self.total_size_label.config(text=f"총 용량: {self.total_attachment_size:.2f} MB / 25.0 MB")

    def reset_form(self):
        if self.is_sending:
            messagebox.showwarning("경고", "이메일 발송 중에는 초기화할 수 없습니다.")
            return
        if not messagebox.askyesno("확인", "모든 데이터를 초기화하시겠습니까?"):
            return
        self.recipients_data.clear()
        self.recipients_tree.delete(*self.recipients_tree.get_children())
        self.attachment_files.clear()
        self.attachment_listbox.delete(0, tk.END)
        self.total_attachment_size = 0
        self.update_total_size_label()
        self.subject.set("파이썬으로 Gmail 보내기")
        self.body_text.delete(1.0, tk.END)
        self.body_text.insert(tk.END, "안녕하세요, {name}님!\n이 메일은 파이썬으로 작성되었습니다.\n감사합니다.")
        self.excel_file_label.config(text="📄 선택된 파일: 없음")
        self.progress_var.set("📋 대기 중...")
        self.progress_bar['value'] = 0
        messagebox.showinfo("완료", "모든 데이터가 초기화되었습니다.")

    def stop_sending(self):
        if self.is_sending:
            self.is_sending = False
            self.progress_var.set("⏹️ 발송 중지 요청...")
            messagebox.showinfo("정보", "발송 중지를 요청했습니다. 현재 작업 완료 후 중단됩니다.")

    def send_email_start(self):
        if not self.recipients_data:
            messagebox.showerror("오류", "수신자를 먼저 로드해주세요.")
            return
        if not self.sender_email.get() or not self.sender_password.get():
            messagebox.showerror("오류", "발신자 정보(Gmail 주소, 앱 비밀번호)를 입력해주세요.")
            return
        if self.is_sending:
            messagebox.showwarning("경고", "이미 발송 중입니다.")
            return
        if not messagebox.askyesno("발송 확인", f"{len(self.recipients_data)}명에게 이메일을 발송하시겠습니까?"):
            return
        self.send_button.config(state='disabled')
        self.is_sending = True
        thread = threading.Thread(target=self._send_email_thread)
        thread.daemon = True
        thread.start()

    def _send_email_thread(self):
        smtp = None
        try:
            self.progress_var.set("🔄 SMTP 서버 연결 중...")
            self.progress_bar['value'] = 0

            smtp = smtplib.SMTP('smtp.gmail.com', 587)
            smtp.starttls()
            smtp.login(self.sender_email.get(), self.sender_password.get())

            total_recipients = len(self.recipients_data)
            success_count = 0
            for idx, recipient in enumerate(self.recipients_data):
                if not self.is_sending:
                    break
                name, email = recipient['name'], recipient['email']
                msg = MIMEMultipart()
                msg['From'] = self.sender_email.get()
                msg['To'] = email
                msg['Subject'] = self.subject.get()

                body_content = self.body_text.get("1.0", tk.END).replace("{name}", name)
                msg.attach(MIMEText(body_content, 'plain'))

                for file_path in self.attachment_files:
                    with open(file_path, 'rb') as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                        msg.attach(part)

                try:
                    smtp.send_message(msg)
                    success_count += 1
                except Exception as e:
                    print(f"[{email}] 전송 오류: {e}")
                self.progress_var.set(f"🚀 발송 진행 중: {idx + 1}/{total_recipients}")
                self.progress_bar['value'] = ((idx + 1) / total_recipients) * 100
                time.sleep(0.2)

            result = f"전송 완료: {success_count}/{total_recipients}"
            if not self.is_sending:
                result = f"사용자 요청으로 중단. 성공: {success_count}/{total_recipients}"

            self.progress_var.set(result)
            messagebox.showinfo("완료", result)

        except Exception as e:
            messagebox.showerror("오류", f"SMTP 연결 또는 전송 중 오류로 중단되었습니다:\n{e}")
        finally:
            if smtp:
                smtp.quit()
            self.send_button.config(state='normal')
            self.is_sending = False


if __name__ == "__main__":
    root = tk.Tk()
    app = GmailSenderGUI(root)
    root.mainloop()