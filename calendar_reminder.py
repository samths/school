"""
calendar_reminder.py 리마인더 캘린더  Ver 1.1_251023
HH:MM -> reminder 내용
"""
import tkinter as tk
from tkinter import messagebox, simpledialog, Scrollbar, Canvas
import calendar
from datetime import datetime
import pandas as pd
import os

# Excel 파일 이름 정의
REMINDER_FILE = 'reminder.xlsx'
ANNIVERSARY_FILE = 'anniversary.xlsx'  # 기념일 파일 추가

class CustomDialog(simpledialog.Dialog):
    """Enter 키 지원 커스텀 다이얼로그"""

    def __init__(self, parent, title, prompt, initialvalue=""):
        self.prompt = prompt
        self.initialvalue = initialvalue
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text=self.prompt).grid(row=0, sticky="w", padx=5, pady=5)
        self.entry = tk.Entry(master, width=30)
        self.entry.grid(row=1, padx=5, pady=5)
        self.entry.insert(0, self.initialvalue)
        self.entry.select_range(0, tk.END)
        self.entry.bind('<Return>', lambda e: self.ok())
        return self.entry

    def apply(self):
        self.result = self.entry.get()

class CalendarReminderApp:
    """
    Python과 Tkinter로 구축된 간단한 캘린더 및 알림 애플리케이션입니다.
    월별 캘린더 보기, 월별 탐색, 날짜별 알림 관리 및 기념일 표시 기능이 포함되어 있습니다.
    """

    def __init__(self, master):
        self.master = master
        master.title("리마인더 캘린더")

        # --- Data Storage ---
        self.reminders = {}
        self.anniversaries = {}  # 기념일 데이터 저장소 추가
        self.load_reminders()
        self.load_anniversaries()  # 기념일 로드

        # --- Date State ---
        now = datetime.now()
        self.current_year = now.year
        self.current_month = now.month
        self.selected_date = None

        # --- Configuration ---
        self.day_names = ["월", "화", "수", "목", "금", "토", "일"]
        self.day_buttons = {}

        # --- 키보드 단축키 바인딩 ---
        self.master.bind('<Control-n>', lambda e: self.add_reminder_dialog())
        self.master.bind('<Control-N>', lambda e: self.add_reminder_dialog())

        # --- UI Setup ---
        self.master.grid_columnconfigure(0, weight=0, minsize=400)  # 달력 폭 400px 고정
        self.master.grid_columnconfigure(1, weight=1)
        self.master.grid_rowconfigure(1, weight=1)

        self.create_header_frame()
        self.create_calendar_frame()
        self.create_reminder_frame()
        self.create_anniversary_frame()  # 기념일 프레임 생성

        self.draw_calendar()

    # --- 기념일 로드/저장 메서드 ---
    def load_anniversaries(self):
        """anniversary.xlsx에서 기념일을 로드합니다. (월, 일, 내용)"""
        self.anniversaries = {}
        if not os.path.exists(ANNIVERSARY_FILE):
            print(f"기념일 파일 '{ANNIVERSARY_FILE}'을(를) 찾을 수 없습니다.")
            return

        try:
            # openpyxl 엔진 명시적으로 사용하여 로드
            df = pd.read_excel(ANNIVERSARY_FILE)

            if df.empty:
                print("기념일 파일이 비었습니다.")
                return

            # '월'과 '일'이 정수형인지 확인하고, 두 자릿수 문자열로 포맷
            for _, row in df.iterrows():
                try:
                    # 열 이름은 '월', '일', '내용'으로 가정
                    month = int(row['월'])
                    day = int(row['일'])
                    desc = str(row['내용']).strip()

                    if 1 <= month <= 12 and 1 <= day <= 31 and desc:
                        # 기념일은 'MM-DD' 형식으로 저장
                        date_key = f"{month:02d}-{day:02d}"

                        if date_key not in self.anniversaries:
                            self.anniversaries[date_key] = []

                        self.anniversaries[date_key].append(desc)
                except (ValueError, KeyError) as e:
                    print(f"기념일 데이터 처리 오류 (행: {row.to_dict()}): {e}")

            print(f"{ANNIVERSARY_FILE}에서 {len(df)}개의 기념일을 성공적으로 로드했습니다.")

        except Exception as e:
            messagebox.showerror("기념일 로딩 오류", f"Excel에서 기념일을 로드하는 데 실패했습니다: {e}")
            self.anniversaries = {}

    # --- 기존 리마인더 로드/저장 메서드 (변경 없음) ---
    def load_reminders(self):
        """해당 파일이 있으면 reminder.xlsx에서 미리 알림을 로드합니다."""
        if not os.path.exists(REMINDER_FILE):
            print(f"알림 파일 '{REMINDER_FILE}'을(를) 찾을 수 없습니다. 빈 알림으로 시작합니다.")
            return

        try:
            df = pd.read_excel(REMINDER_FILE)

            if df.empty:
                print("알림 파일이 비었습니다.")
                return

            for _, row in df.iterrows():
                date_str = str(row['Date'])
                time_str = str(row['Time'])
                desc = str(row['Description'])

                if isinstance(row['Date'], datetime):
                    date_str = row['Date'].strftime("%Y-%m-%d")

                if isinstance(row['Time'], datetime):
                    time_str = row['Time'].strftime("%H:%M")
                elif len(time_str) > 5 and ':' in time_str:
                    try:
                        time_str = datetime.strptime(time_str.split()[-1], '%H:%M:%S').strftime("%H:%M")
                    except ValueError:
                        pass

                if not date_str or not time_str:
                    continue

                if date_str not in self.reminders:
                    self.reminders[date_str] = []

                self.reminders[date_str].append({'time': time_str, 'desc': desc})

            print(f"{REMINDER_FILE}에서 {len(df)}개의 항목을 성공적으로 로드했습니다.")

        except Exception as e:
            messagebox.showerror("로딩 오류", f"Excel에서 알림을 로드하는 데 실패했습니다: {e}")
            self.reminders = {}

    def save_reminders(self):
        """현재 알림을 reminder.xlsx에 저장합니다."""
        all_data = []
        for date_str, reminders_list in self.reminders.items():
            for reminder in reminders_list:
                all_data.append({
                    'Date': date_str,
                    'Time': reminder['time'],
                    'Description': reminder['desc']
                })
        if not all_data:
            df = pd.DataFrame(columns=['Date', 'Time', 'Description'])
        else:
            df = pd.DataFrame(all_data)
        try:
            df.to_excel(REMINDER_FILE, index=False)
            print(f"{len(all_data)}개의 알림을 {REMINDER_FILE}에 성공적으로 저장했습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"Excel에 미리 알림을 저장하지 못했습니다: {e}")

    # --- UI 생성 메서드 ---
    def create_header_frame(self):
        """월 탐색 및 표시를 포함하는 프레임을 만듭니다."""
        header_frame = tk.Frame(self.master, padx=10, pady=10, bg="#2c3e50")
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew")

        prev_button = tk.Button(header_frame, text="< 이전", command=lambda: self.change_month(-1),
                                bg="#3498db", fg="white", font=('맑은 고딕', 10, 'bold'), relief=tk.FLAT, padx=10, pady=5)
        prev_button.pack(side=tk.LEFT, padx=(0, 20))

        self.month_label = tk.Label(header_frame, text="", bg="#2c3e50", fg="white", font=('맑은 고딕', 16, 'bold'))
        self.month_label.pack(side=tk.LEFT, expand=True)

        next_button = tk.Button(header_frame, text="다음 >", command=lambda: self.change_month(1),
                                bg="#3498db", fg="white", font=('맑은 고딕', 10, 'bold'), relief=tk.FLAT, padx=10, pady=5)
        next_button.pack(side=tk.RIGHT, padx=(20, 0))

    def create_calendar_frame(self):
        """고정된 400px 너비로 달력 그리드 표시를 위한 프레임을 만듭니다."""
        self.calendar_frame = tk.Frame(self.master, padx=10, pady=10, bg="#ecf0f1", width=400)
        # 캘린더 프레임은 Row 1, Column 0에 위치하며, pady=(10, 0)으로 아래쪽 여백을 제거하여 기념일 프레임과 가깝게 배치
        self.calendar_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(10, 0))
        self.calendar_frame.grid_propagate(False)  # 크기 고정
        self.calendar_frame.grid_columnconfigure(tuple(range(7)), weight=1)
        for i, day in enumerate(self.day_names):
            day_label = tk.Label(self.calendar_frame, text=day, bg="#bdc3c7", fg="#2c3e50",
                                 font=('맑은 고딕', 11, 'bold'), pady=5)
            # 일요일 (i=6)은 빨간색, 토요일 (i=5)은 파란색
            if i == 5:  # 토요일
                day_label.config(fg="#3498db")
            elif i == 6:  # 일요일
                day_label.config(fg="#e74c3c")
            day_label.grid(row=0, column=i, sticky="ew")

    def create_reminder_frame(self):
        """마우스 휠 지원을 통해 알림 관리 및 표시를 위한 프레임을 만듭니다."""
        self.reminder_frame = tk.Frame(self.master, padx=15, pady=15, bg="#f5f5f5", relief=tk.GROOVE, bd=1)
        self.reminder_frame.grid(row=1, column=1, sticky="nsew", padx=(0, 10), pady=10)
        self.reminder_frame.grid_rowconfigure(1, weight=1)
        self.reminder_frame.grid_columnconfigure(0, weight=1)

        self.reminder_title = tk.Label(self.reminder_frame, text="날자를 선택하시오.", bg="#f5f5f5", fg="#2c3e50", font=('맑은 고딕', 14, 'bold'))
        self.reminder_title.grid(row=0, column=0, sticky="w", pady=(0, 10))

        # 캔버스가 있는 알림 목록 영역
        self.reminder_list_canvas = Canvas(self.reminder_frame, bg="#ffffff", relief=tk.SUNKEN, bd=1)
        self.reminder_list_canvas.grid(row=1, column=0, sticky="nsew")

        # 마우스 휠 스크롤 바인딩
        self.reminder_list_canvas.bind('<MouseWheel>', self._on_mousewheel)
        self.reminder_list_canvas.bind('<Button-4>', self._on_mousewheel)  # Linux
        self.reminder_list_canvas.bind('<Button-5>', self._on_mousewheel)  # Linux

        self.reminder_scrollbar = Scrollbar(self.reminder_list_canvas, orient="vertical", command=self.reminder_list_canvas.yview)
        self.reminder_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.reminder_list_frame = tk.Frame(self.reminder_list_canvas, bg="#ffffff")
        self.reminder_list_canvas.create_window((0, 0), window=self.reminder_list_frame, anchor="nw", width=300)
        self.reminder_list_canvas.configure(yscrollcommand=self.reminder_scrollbar.set)

        self.reminder_list_frame.bind("<Configure>",
                                      lambda e: self.reminder_list_canvas.configure(scrollregion=self.reminder_list_canvas.bbox("all")))

        # 단축키 안내 추가
        shortcut_label = tk.Label(self.reminder_frame, text="단축키: Ctrl+N (새로운 알람)",
                                  bg="#f5f5f5", fg="#7f8c8d", font=('맑은 고딕', 8))
        shortcut_label.grid(row=2, column=0, sticky="w", pady=(5, 5))

        self.add_reminder_button = tk.Button(self.reminder_frame, text="➕ 알림 추가 (Ctrl+N)", command=self.add_reminder_dialog,
                                             bg="#27ae60", fg="white", font=('맑은 고딕', 10, 'bold'), relief=tk.FLAT, padx=10, pady=5, state=tk.DISABLED)
        self.add_reminder_button.grid(row=3, column=0, sticky="ew", pady=(5, 0))

    def create_anniversary_frame(self):
        """캘린더 아래 (row 2, column 0)에 기념일 목록 표시를 위한 프레임을 만듭니다."""
        self.anniversary_frame = tk.Frame(self.master, padx=10, pady=5, bg="#ecf0f1", bd=1, relief=tk.RAISED)
        # 달력 아래 (Row 2, Column 0)에 위치하며, 달력과 동일하게 400px 폭을 유지. pady=(0, 10)으로 위쪽 여백 최소화.
        self.anniversary_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.anniversary_frame.grid_columnconfigure(0, weight=1)

        tk.Label(self.anniversary_frame, text="🗓️ 이번 달의 기념일",
                 bg="#3498db", fg="white", font=('맑은 고딕', 11, 'bold'), anchor="w", padx=5, pady=5).pack(fill=tk.X)

        self.anniversary_list_canvas = Canvas(self.anniversary_frame, bg="#ffffff", height=120)  # 높이 지정
        self.anniversary_list_canvas.pack(fill=tk.X, expand=True)

        self.anniversary_list_frame = tk.Frame(self.anniversary_list_canvas, bg="#ffffff")
        self.anniversary_list_canvas.create_window((0, 0), window=self.anniversary_list_frame, anchor="nw")

        self.anniversary_list_frame.bind("<Configure>",
                                         lambda e: self.anniversary_list_canvas.config(scrollregion=self.anniversary_list_canvas.bbox("all")))

    # --- 달력/화면 업데이트 메서드 ---
    def _on_mousewheel(self, event):
        """마우스 휠 스크롤 이벤트 처리"""
        if event.num == 5 or event.delta < 0:
            self.reminder_list_canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.reminder_list_canvas.yview_scroll(-1, "units")

    def change_month(self, delta):
        """현재 월을 변경하고 달력을 다시 그립니다."""
        new_month = self.current_month + delta
        self.current_year += (new_month - 1) // 12
        self.current_month = (new_month - 1) % 12 + 1

        self.selected_date = None
        self.add_reminder_button.config(state=tk.DISABLED)
        self.update_reminder_display()
        self.draw_calendar()

    def draw_calendar(self):
        """이전 달력을 지우고 현재 월/년도에 대한 격자를 그립니다."""
        month_name = calendar.month_name[self.current_month]
        self.month_label.config(text=f"{month_name} {self.current_year}")

        for widget in self.calendar_frame.winfo_children():
            if int(widget.grid_info()['row']) > 0:
                widget.destroy()
        self.day_buttons.clear()

        # calendar.Calendar(firstweekday=calendar.MONDAY)는 월요일=0, 일요일=6
        month_cal = calendar.Calendar(firstweekday=calendar.MONDAY).monthdays2calendar(
            self.current_year, self.current_month
        )

        current_row = 1
        today = datetime.now().date()
        current_month_str = f"{self.current_month:02d}"

        for week in month_cal:
            for day_index, (day_num, weekday) in enumerate(week):
                if day_num != 0:
                    date_obj = datetime(self.current_year, self.current_month, day_num).date()
                    date_str = date_obj.strftime("%Y-%m-%d")
                    day_month_str = f"{self.current_month:02d}-{day_num:02d}"

                    has_reminder = date_str in self.reminders and len(self.reminders[date_str]) > 0
                    has_anniversary = day_month_str in self.anniversaries and len(self.anniversaries[day_month_str]) > 0

                    # 기본 색상 설정
                    bg_color = "#ffffff"
                    fg_color = "#34495e"  # 평일 기본 글꼴 색상

                    # 1. 주말 색상 적용 (배경색이 바뀌지 않았을 때 글꼴 색상 적용)
                    if weekday == 5:  # 토요일
                        fg_color = "#3498db"  # 파란색
                    elif weekday == 6:  # 일요일
                        fg_color = "#e74c3c"  # 빨간색

                    # 2. 하이라이트 순서 (배경색 변경)
                    if date_obj == today:
                        bg_color = "#f39c12"  # 주황색 (오늘)
                        fg_color = "white"  # 배경색이 바뀌면 글꼴색 흰색
                    elif has_reminder:
                        bg_color = "#e74c3c"  # 빨간색 (알림)
                        fg_color = "white"
                    elif has_anniversary:
                        bg_color = "#3498db"  # 파란색 (기념일)
                        fg_color = "white"

                    # ⚠️ 알림이나 기념일이 주말과 겹칠 경우, 알림/기념일 색상(흰색 글꼴)이 우선 적용됩니다.
                    day_button = tk.Button(self.calendar_frame, text=str(day_num),
                                           command=lambda d=date_obj: self.select_date(d),
                                           bg=bg_color, fg=fg_color, font=('맑은 고딕', 12, 'bold'),
                                           relief=tk.RAISED, borderwidth=1, padx=10, pady=10)
                    day_button.grid(row=current_row, column=day_index, sticky="nsew", padx=2, pady=2)
                    self.day_buttons[date_str] = day_button
                else:
                    empty_label = tk.Label(self.calendar_frame, text="", bg="#ecf0f1")
                    empty_label.grid(row=current_row, column=day_index, sticky="nsew", padx=2, pady=2)
            current_row += 1

        self.update_anniversary_display()

    def select_date(self, date_obj):
        """날짜 버튼을 클릭했을 때의 동작을 처리합니다."""
        date_str = date_obj.strftime("%Y-%m-%d")

        if self.selected_date:
            prev_str = self.selected_date.strftime("%Y-%m-%d")
            if prev_str in self.day_buttons:
                self.redraw_button(prev_str, self.selected_date)

        self.selected_date = date_obj

        if date_str in self.day_buttons:
            self.day_buttons[date_str].config(relief=tk.SUNKEN, bg="#9b59b6")  # 선택된 날짜는 보라색

        self.update_reminder_display()
        self.add_reminder_button.config(state=tk.NORMAL)

    def redraw_button(self, date_str, date_obj):
        """오늘/알림/기념일/주말 상태에 따라 버튼의 모양을 재설정합니다."""
        if date_str in self.day_buttons:

            day_month_str = date_obj.strftime("%m-%d")
            has_reminder = date_str in self.reminders and len(self.reminders[date_str]) > 0
            has_anniversary = day_month_str in self.anniversaries and len(self.anniversaries[day_month_str]) > 0
            today = datetime.now().date()
            weekday = date_obj.weekday()  # 월요일=0, 일요일=6

            bg_color = "#ffffff"
            fg_color = "#34495e"

            # 1. 주말 색상 적용 (글꼴 색상만 변경)
            if weekday == 5:  # 토요일
                fg_color = "#3498db"
            elif weekday == 6:  # 일요일
                fg_color = "#e74c3c"

            # 2. 하이라이트 순서 (배경색 변경 시 글꼴색 흰색으로 덮어쓰기)
            if date_obj == today:
                bg_color = "#f39c12"
                fg_color = "white"
            elif has_reminder:
                bg_color = "#e74c3c"
                fg_color = "white"
            elif has_anniversary:
                bg_color = "#3498db"
                fg_color = "white"

            self.day_buttons[date_str].config(relief=tk.RAISED, bg=bg_color, fg=fg_color)

    def update_anniversary_display(self):
        """사이드바가 아닌, 달력 아래 패널의 기념일 목록을 업데이트합니다."""
        for widget in self.anniversary_list_frame.winfo_children():
            widget.destroy()

        current_month_str = f"{self.current_month:02d}"
        month_anniversaries = []

        # 현재 월의 모든 기념일을 필터링
        for date_key, descs in self.anniversaries.items():
            month, day = date_key.split('-')
            if month == current_month_str:
                for desc in descs:
                    month_anniversaries.append((int(day), desc))

        # 일자별로 정렬
        month_anniversaries.sort(key=lambda x: x[0])

        if not month_anniversaries:
            msg = tk.Label(self.anniversary_list_frame, text="이번 달에는 등록된 기념일이 없습니다.", bg="#ffffff", fg="#7f8c8d")
            msg.pack(fill=tk.X, padx=10, pady=10)
        else:
            for day, desc in month_anniversaries:
                item_label = tk.Label(self.anniversary_list_frame,
                                      text=f"• {day}일: {desc}",
                                      anchor='w', bg="#f0f8ff", font=('맑은 고딕', 10), padx=10, pady=2, relief=tk.FLAT, bd=1)
                item_label.pack(fill=tk.X, pady=1, padx=2)
        self.anniversary_list_frame.update_idletasks()

        # 프레임 너비를 캔버스 너비에 맞추기 위해 캔버스에 윈도우 크기 재설정
        self.anniversary_list_canvas.create_window((0, 0), window=self.anniversary_list_frame, anchor="nw", width=self.anniversary_list_canvas.winfo_width())
        self.anniversary_list_canvas.config(scrollregion=self.anniversary_list_canvas.bbox("all"))

    def update_reminder_display(self):
        """사이드바의 알림 제목과 목록을 업데이트합니다."""
        for widget in self.reminder_list_frame.winfo_children():
            widget.destroy()

        if not self.selected_date:
            self.reminder_title.config(text="날짜를 선택하시오")
            msg = tk.Label(self.reminder_list_frame, text="달력에서 날짜를 클릭하면 해당 날짜의 알림을 관리할 수 있습니다.",
                           wraplength=250, justify=tk.LEFT, bg="#ffffff")
            msg.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            self.add_reminder_button.config(state=tk.DISABLED)
        else:
            date_str = self.selected_date.strftime("%Y-%m-%d")
            display_date = self.selected_date.strftime("%Y년 %m월 %d일 (%A)")
            self.reminder_title.config(text=f"{display_date} 알림")

            reminders_list = self.reminders.get(date_str, [])
            reminders_list.sort(key=lambda r: r['time'])

            if not reminders_list:
                msg = tk.Label(self.reminder_list_frame, text="이 날짜에 설정된 알림이 없습니다.", bg="#ffffff")
                msg.pack(fill=tk.X, padx=10, pady=10)
            else:
                for i, reminder in enumerate(reminders_list):
                    reminder_item_frame = tk.Frame(self.reminder_list_frame, bg="#f5f5f5", padx=5, pady=5, bd=1, relief=tk.RIDGE)
                    reminder_item_frame.pack(fill=tk.X, pady=2, padx=5)
                    reminder_item_frame.grid_columnconfigure(0, weight=1)

                    text_label = tk.Label(reminder_item_frame,
                                          text=f"{i + 1}. {reminder['time']} - {reminder['desc']}",
                                          anchor='w', bg="#f5f5f5", justify=tk.LEFT)
                    text_label.grid(row=0, column=0, sticky="w")

                    edit_btn = tk.Button(reminder_item_frame, text="✎", width=3,
                                         command=lambda idx=i: self.edit_reminder_dialog(idx),
                                         bg="#f1c40f", fg="white", relief=tk.FLAT, font=('맑은 고딕', 8, 'bold'))
                    edit_btn.grid(row=0, column=1, padx=(5, 2))

                    delete_btn = tk.Button(reminder_item_frame, text="✖", width=3,
                                           command=lambda idx=i: self.delete_reminder(idx),
                                           bg="#e74c3c", fg="white", relief=tk.FLAT, font=('맑은 고딕', 8, 'bold'))
                    delete_btn.grid(row=0, column=2)

            self.reminder_list_frame.update_idletasks()
            self.reminder_list_canvas.config(scrollregion=self.reminder_list_canvas.bbox("all"))

            self.add_reminder_button.config(state=tk.NORMAL)

    # --- 알림 관리 메서드 (변경 없음) ---
    def validate_time_format(self, time_input):
        """HH:MM 형식을 검증하는 도우미입니다."""
        try:
            datetime.strptime(time_input, '%H:%M')
            return True
        except ValueError:
            return False

    def add_reminder_dialog(self):
        """Enter 키를 지원하여 새로운 알림 세부 정보를 수집하는 대화 상자를 엽니다."""
        if not self.selected_date:
            messagebox.showwarning("선택된 날짜가 없습니다", "먼저 달력에서 날짜를 선택하세요.")
            return

        date_str = self.selected_date.strftime("%Y-%m-%d")

        # 커스텀 다이얼로그 사용 (Enter 키 지원)
        time_dialog = CustomDialog(self.master, "알림 추가",
                                   "알림 시간을 입력하세요(HH:MM 형식):")
        time_input = time_dialog.result

        if not time_input:
            return

        if not self.validate_time_format(time_input):
            messagebox.showerror("잘못된 시간", "필요한 HH:MM 형식(예: 09:30 또는 14:00)을 사용하세요.")
            return

        desc_dialog = CustomDialog(self.master, "알림 추가", "알림 설명을 입력하세요:")
        desc_input = desc_dialog.result

        if not desc_input:
            return

        new_reminder = {'time': time_input, 'desc': desc_input.strip()}

        if date_str not in self.reminders:
            self.reminders[date_str] = []

        self.reminders[date_str].append(new_reminder)
        self.reminders[date_str].sort(key=lambda r: r['time'])  # 자동 정렬
        self.save_reminders()

        self.update_reminder_display()
        self.redraw_button(date_str, self.selected_date)

        # 개선된 성공 메시지
        display_date = self.selected_date.strftime("%B %d, %Y")
        messagebox.showinfo("✓ 알림 추가",
                            f"알림을 성공적으로 추가했습니다.:\n\n"
                            f"Date: {display_date}\n"
                            f"Time: {time_input}\n"
                            f"Description: {desc_input}")

    def edit_reminder_dialog(self, index):
        """Enter 키를 지원하여 기존 알림을 편집하는 대화 상자를 엽니다."""
        if not self.selected_date:
            return

        date_str = self.selected_date.strftime("%Y-%m-%d")
        reminders_list = self.reminders.get(date_str, [])

        if not (0 <= index < len(reminders_list)):
            messagebox.showerror("오류", "편집을 위한 알림 선택이 잘못되었습니다.")
            return
        original_reminder = reminders_list[index]

        time_dialog = CustomDialog(self.master, "알림 시간 편집",
                                   f"알림 시간 편집:\n{original_reminder['desc']}",
                                   original_reminder['time'])
        new_time = time_dialog.result

        if new_time is None:
            return

        if not self.validate_time_format(new_time):
            messagebox.showerror("잘못된 시간", "필요한 HH:MM 형식(예: 09:30 또는 14:00)을 사용하세요.")
            return

        desc_dialog = CustomDialog(self.master, "알림 설명 편집",
                                   "알림 설명 편집:",
                                   original_reminder['desc'])
        new_desc = desc_dialog.result

        if new_desc is None:
            return

        reminders_list[index]['time'] = new_time
        reminders_list[index]['desc'] = new_desc.strip()
        reminders_list.sort(key=lambda r: r['time'])  # 자동 정렬

        self.save_reminders()

        self.update_reminder_display()
        messagebox.showinfo("성공", "알림이 성공적으로 업데이트 됬습니다")

    def delete_reminder(self, index):
        """선택한 날짜에 대한 0 기반 인덱스를 기준으로 알림을 삭제합니다."""
        if not self.selected_date:
            return

        date_str = self.selected_date.strftime("%Y-%m-%d")
        reminders_list = self.reminders.get(date_str, [])

        if 0 <= index < len(reminders_list):
            deleted_desc = reminders_list[index]['desc']

            del reminders_list[index]
            if not reminders_list:
                del self.reminders[date_str]

            self.save_reminders()

            self.update_reminder_display()
            self.redraw_button(date_str, self.selected_date)
            messagebox.showinfo("성공", f"알림 '{deleted_desc}' 삭제되었습니다.")
        else:
            messagebox.showerror("오류", "삭제할 알림 선택이 잘못되었습니다.")


if __name__ == "__main__":
    if 'pd' not in globals():
        print("경고: pandas 라이브러리가 로드되지 않았습니다. pip install pandas openpyxl 를 실행하여 설치하세요.")

    root = tk.Tk()
    root.option_add("*Font", "Inter 10")
    # 창 높이를 약간 늘려 기념일 목록 공간 확보
    root.geometry("800x680")
    root.resizable(True, True)

    app = CalendarReminderApp(root)
    root.mainloop()