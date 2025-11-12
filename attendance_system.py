"""
attendance_system.py  출석 관리 시스템 Ver 1.0_251026
주요 개선사항: 비밀번호 : pine858!
- 비밀번호 해시화 (보안 강화)
- 체류시간 계산 및 저장 기능 수정
- 실시간 체류시간 표시
- 백업 기능 추가
- 코드 리팩토링 및 에러 처리 강화
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import hashlib
from datetime import datetime
from collections import defaultdict
from shutil import copy2


class AttendanceManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("인원 관리 시스템")
        self.root.geometry("1000x700")

        # 윈도우 아이콘 변경
        try:
            self.root.iconphoto(True, tk.PhotoImage(file='8ball.png'))
        except tk.TclError:
            pass  # 아이콘 없이 시작

        self.data_file = "attendance_data.json"
        self.backup_file = "attendance_data_backup.json"
        self.employees = []
        self.deleted_employees = []
        self.attendance_records = defaultdict(list)
        self.employee_registration_dates = {}
        self.password_hash = ""

        self.load_data()
        self.create_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 실시간 업데이트를 위한 타이머
        self.update_realtime_status()

    def hash_password(self, password):
        """비밀번호 해시화"""
        return hashlib.sha256(password.encode()).hexdigest()

    def verify_password(self, password):
        """비밀번호 검증"""
        return self.hash_password(password) == self.password_hash

    def create_backup(self):
        """데이터 백업 생성"""
        try:
            if os.path.exists(self.data_file):
                copy2(self.data_file, self.backup_file)
                return True
        except Exception as e:
            print(f"백업 실패: {e}")
            return False

    def on_closing(self):
        """프로그램 종료 시 실행"""
        today = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")

        present_employees = self.get_present_employees(today)

        if present_employees:
            for employee in present_employees:
                self.auto_checkout_employee(employee, today, current_time)

            self.save_data()
            messagebox.showinfo("자동 퇴장",
                                f"프로그램 종료로 인해 {len(present_employees)}명이 자동 퇴장 처리되었습니다.\n({current_time})")

        self.root.destroy()

    def get_present_employees(self, date):
        """특정 날짜의 현재 입장중인 인원 목록 반환"""
        present_employees = []
        for employee in self.employees + self.deleted_employees:
            today_records = [record for record in self.attendance_records[employee]
                             if record['date'] == date]
            if today_records and today_records[-1]['check_out'] is None:
                present_employees.append(employee)
        return present_employees

    def auto_checkout_employee(self, employee, date, time):
        """자동 퇴장 처리"""
        today_records = [record for record in self.attendance_records[employee]
                         if record['date'] == date]
        if today_records and today_records[-1]['check_out'] is None:
            today_records[-1]['check_out'] = time
            # 체류시간 계산 및 저장
            self.calculate_and_store_duration(today_records[-1])

    def calculate_and_store_duration(self, record):
        """체류시간 계산하여 레코드에 저장"""
        if record['check_out']:
            try:
                in_time = datetime.strptime(f"{record['date']} {record['check_in']}",
                                            "%Y-%m-%d %H:%M:%S")
                out_time = datetime.strptime(f"{record['date']} {record['check_out']}",
                                             "%Y-%m-%d %H:%M:%S")
                duration = (out_time - in_time).total_seconds() / 3600
                record['duration_hours'] = duration
            except Exception as e:
                print(f"체류시간 계산 오류: {e}")
                record['duration_hours'] = 0
        else:
            record['duration_hours'] = 0

    def format_duration(self, hours):
        """시간을 '시간 분' 형태로 포맷"""
        if hours <= 0:
            return "0시간 0분"
        hours_int = int(hours)
        minutes_int = int((hours - hours_int) * 60)
        return f"{hours_int}시간 {minutes_int}분"

    def parse_duration_str(self, duration_str):
        """'시간 분' 형태의 문자열을 분 단위로 변환"""
        if "시간" in duration_str and "분" in duration_str:
            try:
                hours_part = duration_str.split("시간")[0].strip()
                minutes_part = duration_str.split("시간")[1].split("분")[0].strip()
                return int(hours_part) * 60 + int(minutes_part)
            except:
                return 0
        return -1 if duration_str in ["계산불가", "-"] else 0

    def create_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 인원 관리
        top_frame = ttk.LabelFrame(main_frame, text="인원 관리", padding=10)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        button_frame = ttk.Frame(top_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="인원 추가", command=self.add_employee).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="인원 삭제", command=self.delete_employee).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="백업 생성", command=self.manual_backup).pack(side=tk.LEFT, padx=(0, 5))

        self.total_count_label = ttk.Label(top_frame, text="총 등록 인원: 0명")
        self.total_count_label.pack(anchor=tk.W, pady=(5, 0))

        employee_frame = ttk.Frame(top_frame)
        employee_frame.pack(fill=tk.BOTH, expand=True, padx=(10, 0))

        self.employee_buttons = []
        self.create_employee_buttons(employee_frame)

        # 입장 체크
        middle_frame = ttk.LabelFrame(main_frame, text="입장 체크", padding=10)
        middle_frame.pack(fill=tk.X, pady=(0, 10))

        summary_frame = ttk.Frame(middle_frame)
        summary_frame.pack(fill=tk.X, pady=(0, 10))

        self.today_summary_label = ttk.Label(summary_frame, text="금일 입장 현황: 0명")
        self.today_summary_label.pack(side=tk.LEFT)

        self.present_summary_label = ttk.Label(summary_frame, text="현재 입장중: 0명")
        self.present_summary_label.pack(side=tk.LEFT, padx=(20, 0))

        self.today_frame = ttk.Frame(middle_frame)
        self.today_frame.pack(fill=tk.X)

        ttk.Label(self.today_frame, text="상세 현황:").pack(side=tk.LEFT)

        self.today_status_text = tk.Text(self.today_frame, height=3, wrap=tk.WORD, width=80)
        self.today_status_text.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)

        # 통계
        bottom_frame = ttk.LabelFrame(main_frame, text="통계", padding=10)
        bottom_frame.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(bottom_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        daily_frame = ttk.Frame(notebook)
        notebook.add(daily_frame, text="일별 통계")

        monthly_frame = ttk.Frame(notebook)
        notebook.add(monthly_frame, text="월별 통계")

        yearly_frame = ttk.Frame(notebook)
        notebook.add(yearly_frame, text="연별 통계")

        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="설정")

        self.create_daily_tab(daily_frame)
        self.create_monthly_tab(monthly_frame)
        self.create_yearly_tab(yearly_frame)
        self.create_settings_tab(settings_frame)

        notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self.refresh_employee_buttons()
        self.update_today_status()

    def manual_backup(self):
        """수동 백업 실행"""
        if self.create_backup():
            messagebox.showinfo("백업 완료", "데이터가 성공적으로 백업되었습니다.")
        else:
            messagebox.showerror("백업 실패", "백업 중 오류가 발생했습니다.")

    def update_realtime_status(self):
        """실시간 상태 업데이트 (1분마다)"""
        self.update_today_status()
        self.root.after(60000, self.update_realtime_status)  # 60초마다 업데이트

    def on_tab_changed(self, event):
        """탭 변경 시 설정 탭 진입 확인"""
        notebook = event.widget
        current_tab = notebook.tab(notebook.select(), "text")

        if current_tab == "설정":
            if not self.check_password_for_settings():
                notebook.select(0)

    def check_password_for_settings(self):
        """설정 탭 진입을 위한 비밀번호 확인"""
        if not self.password_hash:
            return True

        password = simpledialog.askstring("비밀번호 확인",
                                          "설정 탭에 진입하려면 비밀번호를 입력하세요:", show='*')

        if password and self.verify_password(password):
            return True
        else:
            messagebox.showerror("오류", "비밀번호가 틀렸습니다.")
            return False

    def create_employee_buttons(self, parent):
        self.buttons_frame = ttk.Frame(parent)
        self.buttons_frame.pack(fill=tk.BOTH, expand=True)

    def refresh_employee_buttons(self):
        for widget in self.buttons_frame.winfo_children():
            widget.destroy()
        self.employee_buttons = []

        sorted_employees = sorted(self.employees)

        for i, employee in enumerate(sorted_employees):
            row = i // 6
            col = i % 6

            btn = ttk.Button(
                self.buttons_frame,
                text=employee,
                command=lambda e=employee: self.check_in_out_employee(e),
                width=8
            )
            btn.grid(row=row, column=col, padx=1, pady=1, sticky="ew")
            self.employee_buttons.append(btn)

        for i in range(6):
            self.buttons_frame.grid_columnconfigure(i, weight=1)

        self.total_count_label.config(text=f"총 등록 인원: {len(self.employees)}명")

    def check_in_out_employee(self, employee):
        """입장/퇴장 체크"""
        today = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")

        today_records = [record for record in self.attendance_records[employee]
                         if record['date'] == today]

        if not today_records:
            # 첫 번째 입장
            new_record = {
                'date': today,
                'check_in': current_time,
                'check_out': None,
                'duration_hours': 0
            }
            self.attendance_records[employee].append(new_record)
            messagebox.showinfo("입장 체크", f"{employee}님 입장 체크 완료! ({current_time})")
        else:
            last_record = today_records[-1]

            if last_record['check_out'] is None:
                # 퇴장 체크
                last_record['check_out'] = current_time
                self.calculate_and_store_duration(last_record)
                messagebox.showwarning("퇴장 체크", f"{employee}님 퇴장 체크 완료! ({current_time})")
            else:
                # 재입장 체크
                new_record = {
                    'date': today,
                    'check_in': current_time,
                    'check_out': None,
                    'duration_hours': 0
                }
                self.attendance_records[employee].append(new_record)
                messagebox.showinfo("재입장 체크", f"{employee}님 재입장 체크 완료! ({current_time})")

        self.save_data()
        self.update_today_status()

    def update_today_status(self):
        """오늘의 출근 현황 업데이트"""
        today = datetime.now().strftime("%Y-%m-%d")
        all_attended_employees = []
        present_employees = []

        sorted_employees = sorted(self.employees)

        for employee in sorted_employees:
            today_records = [record for record in self.attendance_records[employee]
                             if record['date'] == today]

            if today_records:
                all_attended_employees.append(employee)
                last_record = today_records[-1]
                if last_record['check_out'] is None:
                    present_employees.append(employee)

        self.today_summary_label.config(text=f"금일 입장 현황: {len(all_attended_employees)}명")
        self.present_summary_label.config(text=f"현재 입장중: {len(present_employees)}명")

        self.today_status_text.delete(1.0, tk.END)

        if present_employees:
            status_lines = []
            for employee in sorted(present_employees):
                # 실시간 체류시간 계산
                today_records = [record for record in self.attendance_records[employee]
                                 if record['date'] == today]
                if today_records:
                    last_record = today_records[-1]
                    if last_record['check_out'] is None:
                        try:
                            in_time = datetime.strptime(f"{today} {last_record['check_in']}",
                                                        "%Y-%m-%d %H:%M:%S")
                            now = datetime.now()
                            duration = (now - in_time).total_seconds() / 3600
                            duration_str = self.format_duration(duration)
                            status_lines.append(f"{employee} ({duration_str})")
                        except:
                            status_lines.append(employee)

            status_text = ", ".join(status_lines)
            self.today_status_text.insert(1.0, status_text)
        else:
            self.today_status_text.insert(1.0, "현재 입장중인 인원이 없습니다.")

    def create_daily_tab(self, parent):
        date_frame = ttk.Frame(parent)
        date_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(date_frame, text="년도:").pack(side=tk.LEFT)

        self.daily_year_var = tk.StringVar(value=str(datetime.now().year))
        self.daily_year_combo = ttk.Combobox(date_frame, textvariable=self.daily_year_var,
                                             values=[str(year) for year in range(2025, 2036)],
                                             width=8, state="readonly")
        self.daily_year_combo.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Label(date_frame, text="월:").pack(side=tk.LEFT)

        self.daily_month_var = tk.StringVar(value=str(datetime.now().month))
        self.daily_month_combo = ttk.Combobox(date_frame, textvariable=self.daily_month_var,
                                              values=[str(month) for month in range(1, 13)],
                                              width=5, state="readonly")
        self.daily_month_combo.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Label(date_frame, text="일:").pack(side=tk.LEFT)

        self.daily_day_var = tk.StringVar(value=str(datetime.now().day))
        self.daily_day_combo = ttk.Combobox(date_frame, textvariable=self.daily_day_var,
                                            values=[str(day) for day in range(1, 32)],
                                            width=5, state="readonly")
        self.daily_day_combo.pack(side=tk.LEFT, padx=(5, 10))

        self.daily_month_combo.bind('<<ComboboxSelected>>', self.update_daily_days)

        ttk.Button(date_frame, text="조회", command=self.show_daily_stats).pack(side=tk.LEFT)

        columns = ("이름", "입장시간", "퇴장시간", "체류시간")
        self.daily_tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)

        for col in columns:
            self.daily_tree.heading(col, text=col)
            self.daily_tree.column(col, width=150)

        self.daily_tree.pack(fill=tk.BOTH, expand=True)
        self.daily_tree.bind("<Double-1>", self.show_daily_detail)

    def update_daily_days(self, event=None):
        """월 변경 시 일 수 업데이트"""
        try:
            year = int(self.daily_year_var.get())
            month = int(self.daily_month_var.get())

            if month in [1, 3, 5, 7, 8, 10, 12]:
                max_day = 31
            elif month in [4, 6, 9, 11]:
                max_day = 30
            else:
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    max_day = 29
                else:
                    max_day = 28

            self.daily_day_combo['values'] = [str(day) for day in range(1, max_day + 1)]

            current_day = int(self.daily_day_var.get())
            if current_day > max_day:
                self.daily_day_var.set(str(max_day))

        except ValueError:
            pass

    def show_daily_detail(self, event):
        """일별 통계 상세 정보 표시"""
        selection = self.daily_tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.daily_tree.item(item, "values")
        employee_name = values[0].replace(" (삭제됨)", "")
        selected_year = self.daily_year_var.get()
        selected_month = self.daily_month_var.get()
        selected_day = self.daily_day_var.get()
        date = f"{selected_year}-{selected_month.zfill(2)}-{selected_day.zfill(2)}"

        daily_records = [record for record in self.attendance_records[employee_name]
                         if record['date'] == date]

        if not daily_records:
            messagebox.showinfo("상세 정보", f"{employee_name}님은 {date}에 출근 기록이 없습니다.")
            return

        detail_window = tk.Toplevel(self.root)
        detail_window.title(f"{employee_name}님 - {date} 상세 기록")
        detail_window.geometry("500x500")
        detail_window.transient(self.root)
        detail_window.grab_set()

        main_frame = ttk.Frame(detail_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text=f"{employee_name}님의 {date} 입장 기록",
                                font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))

        registration_date = self.employee_registration_dates.get(employee_name, "정보 없음")
        reg_date_label = ttk.Label(main_frame, text=f"등록일: {registration_date}",
                                   font=("Arial", 10))
        reg_date_label.pack(pady=(0, 20))

        records_text = tk.Text(main_frame, wrap=tk.WORD, height=15)
        records_text.pack(fill=tk.BOTH, expand=True)

        content = ""
        total_hours = 0

        for i, record in enumerate(daily_records, 1):
            content += f"【{i}번째 입장】\n"
            content += f"입장시간: {record['check_in']}\n"

            if record['check_out']:
                content += f"퇴장시간: {record['check_out']}\n"

                # 저장된 duration_hours 사용 또는 재계산
                if 'duration_hours' in record and record['duration_hours'] > 0:
                    duration_hours = record['duration_hours']
                else:
                    try:
                        in_time = datetime.strptime(f"{date} {record['check_in']}",
                                                    "%Y-%m-%d %H:%M:%S")
                        out_time = datetime.strptime(f"{date} {record['check_out']}",
                                                     "%Y-%m-%d %H:%M:%S")
                        duration_hours = (out_time - in_time).total_seconds() / 3600
                        record['duration_hours'] = duration_hours
                    except:
                        duration_hours = 0

                total_hours += duration_hours
                content += f"체류시간: {self.format_duration(duration_hours)}\n"
            else:
                content += f"퇴장시간: 미퇴장 (현재 입장중)\n"
                content += f"체류시간: 계산불가\n"

            content += "\n" + "=" * 30 + "\n\n"

        if total_hours > 0:
            content += f"【총 체류시간】\n{self.format_duration(total_hours)}\n"

        records_text.insert(1.0, content)
        records_text.config(state=tk.DISABLED)

        ttk.Button(main_frame, text="닫기", command=detail_window.destroy).pack(pady=(10, 0))

    def create_monthly_tab(self, parent):
        month_frame = ttk.Frame(parent)
        month_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(month_frame, text="년도:").pack(side=tk.LEFT)

        self.monthly_year_var = tk.StringVar(value=str(datetime.now().year))
        self.monthly_year_combo = ttk.Combobox(month_frame, textvariable=self.monthly_year_var,
                                               values=[str(year) for year in range(2025, 2036)],
                                               width=8, state="readonly")
        self.monthly_year_combo.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Label(month_frame, text="월:").pack(side=tk.LEFT)

        self.monthly_month_var = tk.StringVar(value=str(datetime.now().month))
        self.monthly_month_combo = ttk.Combobox(month_frame, textvariable=self.monthly_month_var,
                                                values=[str(month) for month in range(1, 13)],
                                                width=5, state="readonly")
        self.monthly_month_combo.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Button(month_frame, text="조회", command=self.show_monthly_stats).pack(side=tk.LEFT)

        columns = ("이름", "입장일수", "총 체류시간", "평균 체류시간")
        self.monthly_tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)

        for col in columns:
            self.monthly_tree.heading(col, text=col)
            self.monthly_tree.column(col, width=150)

        self.monthly_tree.pack(fill=tk.BOTH, expand=True)

    def create_yearly_tab(self, parent):
        year_frame = ttk.Frame(parent)
        year_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(year_frame, text="년도:").pack(side=tk.LEFT)

        self.yearly_year_var = tk.StringVar(value=str(datetime.now().year))
        self.yearly_year_combo = ttk.Combobox(year_frame, textvariable=self.yearly_year_var,
                                              values=[str(year) for year in range(2025, 2036)],
                                              width=8, state="readonly")
        self.yearly_year_combo.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Button(year_frame, text="조회", command=self.show_yearly_stats).pack(side=tk.LEFT)

        columns = ("이름", "총 입장일수", "총 체류시간", "평균 체류시간")
        self.yearly_tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)

        for col in columns:
            self.yearly_tree.heading(col, text=col)
            self.yearly_tree.column(col, width=150)

        self.yearly_tree.pack(fill=tk.BOTH, expand=True)

    def create_settings_tab(self, parent):
        """설정 탭 생성"""
        settings_frame = ttk.Frame(parent, padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True)

        password_frame = ttk.LabelFrame(settings_frame, text="비밀번호 설정", padding=10)
        password_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(password_frame, text="새 비밀번호:").pack(side=tk.LEFT, padx=(0, 5))
        self.password_entry = ttk.Entry(password_frame, show='*', width=20)
        self.password_entry.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(password_frame, text="비밀번호 설정",
                   command=self.set_password).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(password_frame, text="비밀번호 해제",
                   command=self.remove_password).pack(side=tk.LEFT, padx=(0, 5))

        self.password_status_label = ttk.Label(password_frame, text="현재 상태: 비밀번호 미설정")
        self.password_status_label.pack(side=tk.LEFT, padx=(20, 0))

        self.update_password_status()

    def update_password_status(self):
        """비밀번호 상태 업데이트"""
        if self.password_hash:
            self.password_status_label.config(text="현재 상태: 비밀번호 설정됨")
        else:
            self.password_status_label.config(text="현재 상태: 비밀번호 미설정")

    def set_password(self):
        """비밀번호 설정"""
        new_password = self.password_entry.get().strip()

        if not new_password:
            messagebox.showwarning("경고", "비밀번호를 입력하세요.")
            return

        if self.password_hash:
            current_password = simpledialog.askstring("기존 비밀번호 확인",
                                                      "기존 비밀번호를 입력하세요:", show='*')
            if not current_password or not self.verify_password(current_password):
                messagebox.showerror("오류", "기존 비밀번호가 틀렸습니다.")
                return

        self.password_hash = self.hash_password(new_password)
        self.password_entry.delete(0, tk.END)
        self.update_password_status()
        self.save_data()
        messagebox.showinfo("성공", "비밀번호가 설정되었습니다.")

    def remove_password(self):
        """비밀번호 해제"""
        if not self.password_hash:
            messagebox.showinfo("알림", "설정된 비밀번호가 없습니다.")
            return

        current_password = simpledialog.askstring("비밀번호 확인",
                                                  "현재 비밀번호를 입력하세요:", show='*')
        if not current_password or not self.verify_password(current_password):
            messagebox.showerror("오류", "비밀번호가 틀렸습니다.")
            return

        self.password_hash = ""
        self.update_password_status()
        self.save_data()
        messagebox.showinfo("성공", "비밀번호가 해제되었습니다.")

    def add_employee(self):
        """인원 추가"""
        name = simpledialog.askstring("인원 추가", "추가할 이름을 입력하세요:")
        if name and name.strip():
            if name not in self.employees:
                self.employees.append(name)
                self.employee_registration_dates[name] = datetime.now().strftime("%Y-%m-%d")
                self.save_data()
                self.refresh_employee_buttons()
                messagebox.showinfo("성공", f"{name}이(가) 추가되었습니다.")
            else:
                messagebox.showwarning("경고", "이미 존재하는 이름입니다.")

    def delete_employee(self):
        """인원 삭제"""
        if not self.employees:
            messagebox.showwarning("경고", "삭제할 인원이 없습니다.")
            return

        if self.password_hash:
            password = simpledialog.askstring("비밀번호 확인",
                                              "인원 삭제를 위해 비밀번호를 입력하세요:", show='*')
            if not password or not self.verify_password(password):
                messagebox.showerror("오류", "비밀번호가 틀렸습니다.")
                return

        delete_window = tk.Toplevel(self.root)
        delete_window.title("인원 삭제")
        delete_window.geometry("300x350")
        delete_window.transient(self.root)
        delete_window.grab_set()

        main_frame = ttk.Frame(delete_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(main_frame, text="삭제할 인원을 선택하세요:").pack(pady=10)

        listbox_frame = ttk.Frame(main_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=listbox.yview)

        sorted_employees = sorted(self.employees)
        for employee in sorted_employees:
            listbox.insert(tk.END, employee)

        def confirm_delete():
            selection = listbox.curselection()
            if selection:
                employee = sorted_employees[selection[0]]
                if messagebox.askyesno("확인",
                                       f"{employee}을(를) 삭제하시겠습니까?\n\n삭제된 인원의 데이터는 유지되며,\n통계에서 빨간색으로 표시됩니다."):
                    self.employees.remove(employee)
                    self.deleted_employees.append(employee)
                    self.save_data()
                    self.refresh_employee_buttons()
                    self.update_today_status()
                    delete_window.destroy()
                    messagebox.showinfo("성공", f"{employee}이(가) 삭제되었습니다.")
            else:
                messagebox.showwarning("경고", "삭제할 인원을 선택하세요.")

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="삭제", command=confirm_delete).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="취소", command=delete_window.destroy).pack(side=tk.LEFT, padx=5)

    def show_daily_stats(self):
        """일별 통계 표시"""
        for item in self.daily_tree.get_children():
            self.daily_tree.delete(item)

        selected_year = self.daily_year_var.get()
        selected_month = self.daily_month_var.get()
        selected_day = self.daily_day_var.get()
        date = f"{selected_year}-{selected_month.zfill(2)}-{selected_day.zfill(2)}"

        sorted_employees = sorted(self.employees + self.deleted_employees)

        for employee in sorted_employees:
            daily_records = [record for record in self.attendance_records[employee]
                             if record['date'] == date]

            if daily_records:
                all_check_ins = []
                all_check_outs = []
                total_hours = 0

                for record in daily_records:
                    all_check_ins.append(record['check_in'])
                    if record['check_out']:
                        all_check_outs.append(record['check_out'])

                        # 저장된 duration_hours 사용 또는 재계산
                        if 'duration_hours' in record and record['duration_hours'] > 0:
                            total_hours += record['duration_hours']
                        else:
                            try:
                                in_time = datetime.strptime(f"{date} {record['check_in']}",
                                                            "%Y-%m-%d %H:%M:%S")
                                out_time = datetime.strptime(f"{date} {record['check_out']}",
                                                             "%Y-%m-%d %H:%M:%S")
                                duration_hours = (out_time - in_time).total_seconds() / 3600
                                record['duration_hours'] = duration_hours
                                total_hours += duration_hours
                            except:
                                pass

                check_in_str = ", ".join(all_check_ins)
                check_out_str = ", ".join(all_check_outs) if all_check_outs else "미퇴장"
                duration_str = self.format_duration(total_hours) if total_hours > 0 else "계산불가"

                tag = "deleted" if employee in self.deleted_employees else ""
                display_name = f"{employee} (삭제됨)" if employee in self.deleted_employees else employee
                self.daily_tree.insert("", tk.END,
                                       values=(display_name, check_in_str, check_out_str, duration_str),
                                       tags=(tag,))
            else:
                tag = "deleted" if employee in self.deleted_employees else ""
                display_name = f"{employee} (삭제됨)" if employee in self.deleted_employees else employee
                self.daily_tree.insert("", tk.END,
                                       values=(display_name, "미입장", "-", "-"),
                                       tags=(tag,))

        self.daily_tree.tag_configure("deleted", foreground="red")
        self.sort_tree_by_column(self.daily_tree, 3)

    def show_monthly_stats(self):
        """월별 통계 표시"""
        for item in self.monthly_tree.get_children():
            self.monthly_tree.delete(item)

        selected_year = self.monthly_year_var.get()
        selected_month = self.monthly_month_var.get()
        month = f"{selected_year}-{selected_month.zfill(2)}"

        sorted_employees = sorted(self.employees + self.deleted_employees)

        for employee in sorted_employees:
            monthly_records = [record for record in self.attendance_records[employee]
                               if record['date'].startswith(month)]

            if monthly_records:
                work_days = len(set(record['date'] for record in monthly_records))
                total_hours = 0

                for record in monthly_records:
                    if record['check_out']:
                        if 'duration_hours' in record and record['duration_hours'] > 0:
                            total_hours += record['duration_hours']
                        else:
                            try:
                                in_time = datetime.strptime(f"{record['date']} {record['check_in']}",
                                                            "%Y-%m-%d %H:%M:%S")
                                out_time = datetime.strptime(f"{record['date']} {record['check_out']}",
                                                             "%Y-%m-%d %H:%M:%S")
                                duration_hours = (out_time - in_time).total_seconds() / 3600
                                record['duration_hours'] = duration_hours
                                total_hours += duration_hours
                            except:
                                pass

                avg_hours = total_hours / work_days if work_days > 0 else 0

                tag = "deleted" if employee in self.deleted_employees else ""
                display_name = f"{employee} (삭제됨)" if employee in self.deleted_employees else employee
                self.monthly_tree.insert("", tk.END, values=(
                    display_name,
                    f"{work_days}일",
                    self.format_duration(total_hours),
                    self.format_duration(avg_hours)
                ), tags=(tag,))
            else:
                tag = "deleted" if employee in self.deleted_employees else ""
                display_name = f"{employee} (삭제됨)" if employee in self.deleted_employees else employee
                self.monthly_tree.insert("", tk.END,
                                         values=(display_name, "0일", "0시간 0분", "0시간 0분"),
                                         tags=(tag,))

        self.monthly_tree.tag_configure("deleted", foreground="red")
        self.sort_tree_multi_column(self.monthly_tree, [1, 2, 3])

    def show_yearly_stats(self):
        """연별 통계 표시"""
        for item in self.yearly_tree.get_children():
            self.yearly_tree.delete(item)

        selected_year = self.yearly_year_var.get()

        sorted_employees = sorted(self.employees + self.deleted_employees)

        for employee in sorted_employees:
            yearly_records = [record for record in self.attendance_records[employee]
                              if record['date'].startswith(selected_year)]

            if yearly_records:
                work_days = len(set(record['date'] for record in yearly_records))
                total_hours = 0

                for record in yearly_records:
                    if record['check_out']:
                        if 'duration_hours' in record and record['duration_hours'] > 0:
                            total_hours += record['duration_hours']
                        else:
                            try:
                                in_time = datetime.strptime(f"{record['date']} {record['check_in']}",
                                                            "%Y-%m-%d %H:%M:%S")
                                out_time = datetime.strptime(f"{record['date']} {record['check_out']}",
                                                             "%Y-%m-%d %H:%M:%S")
                                duration_hours = (out_time - in_time).total_seconds() / 3600
                                record['duration_hours'] = duration_hours
                                total_hours += duration_hours
                            except:
                                pass

                avg_hours = total_hours / work_days if work_days > 0 else 0

                tag = "deleted" if employee in self.deleted_employees else ""
                display_name = f"{employee} (삭제됨)" if employee in self.deleted_employees else employee
                self.yearly_tree.insert("", tk.END, values=(
                    display_name,
                    f"{work_days}일",
                    self.format_duration(total_hours),
                    self.format_duration(avg_hours)
                ), tags=(tag,))
            else:
                tag = "deleted" if employee in self.deleted_employees else ""
                display_name = f"{employee} (삭제됨)" if employee in self.deleted_employees else employee
                self.yearly_tree.insert("", tk.END,
                                        values=(display_name, "0일", "0시간 0분", "0시간 0분"),
                                        tags=(tag,))

        self.yearly_tree.tag_configure("deleted", foreground="red")
        self.sort_tree_multi_column(self.yearly_tree, [1, 2, 3])

    def sort_tree_by_column(self, tree, col_index):
        """단일 컬럼 기준으로 트리뷰 정렬"""
        items = []
        for item in tree.get_children():
            values = tree.item(item, "values")
            tags = tree.item(item, "tags")
            sort_value = self.parse_duration_str(values[col_index])
            items.append((sort_value, item, values, tags))

        items.sort(key=lambda x: x[0], reverse=True)

        for item in tree.get_children():
            tree.delete(item)

        for _, _, values, tags in items:
            tree.insert("", tk.END, values=values, tags=tags)

    def sort_tree_multi_column(self, tree, col_indices):
        """다중 컬럼 기준으로 트리뷰 정렬"""
        items = []
        for item in tree.get_children():
            values = tree.item(item, "values")
            tags = tree.item(item, "tags")

            sort_values = []
            for col_idx in col_indices:
                if col_idx == 1:  # 일수
                    sort_values.append(int(values[col_idx].replace("일", "")))
                else:  # 시간
                    sort_values.append(self.parse_duration_str(values[col_idx]))

            items.append((tuple(sort_values), item, values, tags))

        items.sort(key=lambda x: x[0], reverse=True)

        for item in tree.get_children():
            tree.delete(item)

        for _, _, values, tags in items:
            tree.insert("", tk.END, values=values, tags=tags)

    def save_data(self):
        """데이터 저장"""
        try:
            # 백업 생성
            self.create_backup()

            data = {
                'employees': self.employees,
                'deleted_employees': self.deleted_employees,
                'attendance_records': dict(self.attendance_records),
                'employee_registration_dates': self.employee_registration_dates,
                'password_hash': self.password_hash
            }

            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("저장 오류", f"데이터 저장 중 오류가 발생했습니다: {e}")

    def load_data(self):
        """데이터 로드"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.employees = data.get('employees', [])
                    self.deleted_employees = data.get('deleted_employees', [])
                    self.attendance_records = defaultdict(list, data.get('attendance_records', {}))
                    self.employee_registration_dates = data.get('employee_registration_dates', {})

                    # 이전 버전 호환성: password를 password_hash로 변환
                    if 'password' in data and data['password']:
                        self.password_hash = self.hash_password(data['password'])
                    else:
                        self.password_hash = data.get('password_hash', '')

                    # 기존 레코드에 duration_hours 추가 (없는 경우)
                    for employee in self.attendance_records:
                        for record in self.attendance_records[employee]:
                            if 'duration_hours' not in record:
                                self.calculate_and_store_duration(record)

            except Exception as e:
                messagebox.showerror("로드 오류", f"데이터 로드 중 오류가 발생했습니다: {e}\n백업 파일을 확인하세요.")
                self.initialize_empty_data()
        else:
            self.initialize_empty_data()

    def initialize_empty_data(self):
        """빈 데이터로 초기화"""
        self.employees = []
        self.deleted_employees = []
        self.attendance_records = defaultdict(list)
        self.employee_registration_dates = {}
        self.password_hash = ''

    def run(self):
        """프로그램 실행"""
        self.root.mainloop()


if __name__ == "__main__":
    app = AttendanceManager()
    app.run()