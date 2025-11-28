"""
booklibrary.py 도서관 도서 대출 관리 Ver 1.0_251114
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
from datetime import datetime, timedelta
import pandas as pd  # pandas 모듈 추가

DB_FILE = "libraryit.db"  # 통합 데이터베이스 파일
FINE_PER_DAY = 500  # 연체료 (원/일)


# --- SQLite3 데이터베이스 함수 ---

def get_db_connection():
    return sqlite3.connect(DB_FILE)


def init_db():
    """데이터베이스 파일이 없으면 생성하고 테이블을 초기화하며, 기존 DB의 구조를 확인하여 업데이트합니다."""
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. BOOKS 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS books (
            book_id TEXT PRIMARY KEY,
            title TEXT NOT NULL,
            author TEXT NOT NULL,
            category TEXT,
            copies INTEGER NOT NULL,
            available INTEGER NOT NULL,
            added_date TEXT
        )
    """)

    # 2. STUDENTS 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            student_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            major TEXT,
            phone TEXT,
            enroll_date TEXT
        )
    """)

    # 3. TRANSACTIONS 테이블
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            transaction_id TEXT PRIMARY KEY,
            student_id TEXT,
            book_id TEXT,
            issue_date TEXT,
            due_date TEXT,
            return_date TEXT,
            status TEXT, -- '대출' or '반납'
            fine INTEGER,
            FOREIGN KEY(student_id) REFERENCES students(student_id),
            FOREIGN KEY(book_id) REFERENCES books(book_id)
        )
    """)

    # --- DB 구조 업데이트 (Migration) ---

    # 1. 'available' 컬럼이 없는 경우 추가 (이전 오류 해결 로직)
    try:
        cursor.execute("SELECT available FROM books LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE books ADD COLUMN available INTEGER NOT NULL DEFAULT 0")
        conn.commit()

    # 2. 'copies' 컬럼이 없는 경우 추가
    try:
        cursor.execute("SELECT copies FROM books LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE books ADD COLUMN copies INTEGER NOT NULL DEFAULT 0")
        conn.commit()

    # 3. 'enroll_date' 컬럼이 없는 경우 추가
    try:
        cursor.execute("SELECT enroll_date FROM students LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE students ADD COLUMN enroll_date TEXT")
        conn.commit()

    conn.commit()
    conn.close()


def generate_id(prefix, table_name):
    """DB에서 마지막 ID를 조회하여 새로운 ID를 생성합니다. (컬럼 이름 오류 수정됨)"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # 테이블 이름에 따라 정확한 ID 컬럼 이름 정의
    if table_name == "books":
        id_column = "book_id"
    elif table_name == "students":
        id_column = "student_id"
    elif table_name == "transactions":
        id_column = "transaction_id"
    else:
        id_column = f"{table_name[:-1]}_id"

        # ID가 존재하는지 확인하고 최신 ID를 가져옴
    cursor.execute(f"SELECT {id_column} FROM {table_name} WHERE {id_column} LIKE '{prefix}%%' ORDER BY {id_column} DESC LIMIT 1")

    last_id_row = cursor.fetchone()

    if last_id_row:
        last_id = last_id_row[0]
        # 예: BK001 -> 1
        # 숫자 추출 시 오류 방지를 위해 prefix 이후 문자열이 모두 숫자인지 확인
        num_part = last_id.replace(prefix, "")
        if num_part and num_part.isdigit():
            num = int(num_part)
            new_id = num + 1
        else:
            # 예상치 못한 ID 형식일 경우 1부터 다시 시작
            new_id = 1
    else:
        new_id = 1

    conn.close()
    return f"{prefix}{str(new_id).zfill(4)}"


# --- Custom Message Box Functions (한국어 텍스트) ---

def show_info(title, message, parent=None):
    """Custom info dialog that stays on top"""
    if parent:
        parent.wm_attributes("-topmost", True)
        parent.lift()
        parent.focus_force()
        result = messagebox.showinfo(title, message, parent=parent)
        parent.wm_attributes("-topmost", False)
        return result
    else:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", True)
        result = messagebox.showinfo(title, message, parent=root)
        root.destroy()
        return result


def show_error(title, message, parent=None):
    """Custom error dialog that stays on top"""
    if parent:
        parent.wm_attributes("-topmost", True)
        parent.lift()
        parent.focus_force()
        result = messagebox.showerror(title, message, parent=parent)
        parent.wm_attributes("-topmost", False)
        return result
    else:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", True)
        result = messagebox.showerror(title, message, parent=root)
        root.destroy()
        return result


def show_warning(title, message, parent=None):
    """Custom warning dialog that stays on top"""
    if parent:
        parent.wm_attributes("-topmost", True)
        parent.lift()
        parent.focus_force()
        result = messagebox.showwarning(title, message, parent=parent)
        parent.wm_attributes("-topmost", False)
        return result
    else:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", True)
        result = messagebox.showwarning(title, message, parent=root)
        root.destroy()
        return result


def ask_yesno(title, message, parent=None):
    """Custom yes/no dialog that stays on top"""
    if parent:
        parent.wm_attributes("-topmost", True)
        parent.lift()
        parent.focus_force()
        result = messagebox.askyesno(title, message, parent=parent)
        parent.wm_attributes("-topmost", False)
        return result
    else:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", True)
        result = messagebox.askyesno(title, message, parent=root)
        root.destroy()
        return result


# Searchable Combobox Class (변경 없음)
class SearchableCombobox(ttk.Frame):
    def __init__(self, parent, textvariable=None, values=None, **kwargs):
        super().__init__(parent)

        self.textvariable = textvariable or tk.StringVar()
        self.values = values or []
        self.filtered_values = self.values.copy()

        # Create entry and listbox
        self.entry = ttk.Entry(self, textvariable=self.textvariable, **kwargs)
        self.entry.pack(fill="x")

        # Listbox for dropdown
        self.listbox_frame = ttk.Frame(self)
        self.listbox = tk.Listbox(self.listbox_frame, height=6)
        self.listbox_scrollbar = ttk.Scrollbar(self.listbox_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=self.listbox_scrollbar.set)

        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox_scrollbar.pack(side="right", fill="y")

        # Bind events
        self.entry.bind("<KeyRelease>", self.on_key_release)
        self.entry.bind("<Button-1>", self.show_dropdown)
        self.entry.bind("<FocusIn>", self.show_dropdown)
        self.listbox.bind("<ButtonRelease-1>", self.on_select)
        self.listbox.bind("<Return>", self.on_select)

        # Initially hide dropdown
        self.dropdown_visible = False

        self.update_listbox()

    def show_dropdown(self, event=None):
        if not self.dropdown_visible and self.filtered_values:
            self.listbox_frame.pack(fill="x", pady=(2, 0))
            self.dropdown_visible = True
            self.update_listbox()

    def hide_dropdown(self):
        if self.dropdown_visible:
            self.listbox_frame.pack_forget()
            self.dropdown_visible = False

    def on_key_release(self, event):
        if event.keysym in ['Up', 'Down', 'Return']:
            return

        search_text = self.textvariable.get().lower()
        self.filtered_values = [val for val in self.values if search_text in val.lower()]

        if self.filtered_values:
            self.show_dropdown()
            self.update_listbox()
        else:
            self.hide_dropdown()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for value in self.filtered_values:
            self.listbox.insert(tk.END, value)

    def on_select(self, event):
        selection = self.listbox.curselection()
        if selection:
            selected_value = self.listbox.get(selection[0])
            self.textvariable.set(selected_value)
            self.hide_dropdown()

    def set_values(self, values):
        self.values = values
        self.filtered_values = values.copy()
        self.update_listbox()

    def get(self):
        return self.textvariable.get()

    def set(self, value):
        self.textvariable.set(value)


class IssueReturnTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent.notebook)
        self.parent = parent

        # Main container
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=5)

        # Title (Style 적용)
        ttk.Label(main_frame, text="도서 대출 / 반납", style="Title.TLabel").pack(pady=(0, 5))

        # Issue Book Section (대출)
        issue_frame = ttk.LabelFrame(main_frame, text="도서 대출", padding=5)
        issue_frame.pack(fill="x", pady=(0, 10))

        # Student ID with Search
        ttk.Label(issue_frame, text="학생 ID:").grid(row=0, column=0, sticky="w", pady=5)
        self.student_id_var = tk.StringVar()
        self.student_search = SearchableCombobox(issue_frame, textvariable=self.student_id_var, width=35)
        self.student_search.grid(row=0, column=1, padx=(10, 0), pady=5, sticky="ew")

        # Book ID with Search
        ttk.Label(issue_frame, text="도서 ID:").grid(row=1, column=0, sticky="w", pady=5)
        self.book_id_var = tk.StringVar()
        self.book_search = SearchableCombobox(issue_frame, textvariable=self.book_id_var, width=35)
        self.book_search.grid(row=1, column=1, padx=(10, 0), pady=5, sticky="ew")

        # Configure grid weights
        issue_frame.grid_columnconfigure(1, weight=1)

        # Issue Button
        ttk.Button(issue_frame, text="도서 대출", command=self.issue_book).grid(row=2, column=0, columnspan=2, pady=5)

        # Return Book Section (반납)
        return_frame = ttk.LabelFrame(main_frame, text="도서 반납", padding=5)
        return_frame.pack(fill="x", pady=10)

        # Transaction ID for return with Search
        ttk.Label(return_frame, text="거래 ID:").grid(row=0, column=0, sticky="w", pady=5)
        self.transaction_id_var = tk.StringVar()
        self.transaction_search = SearchableCombobox(return_frame, textvariable=self.transaction_id_var, width=35)
        self.transaction_search.grid(row=0, column=1, padx=(10, 0), pady=5, sticky="ew")

        # Configure grid weights
        return_frame.grid_columnconfigure(1, weight=1)

        # Return Button
        ttk.Button(return_frame, text="도서 반납", command=self.return_book).grid(row=1, column=0, columnspan=2, pady=10)

        # Current Transactions (Style 적용)
        ttk.Label(main_frame, text="현재 대출 중인 도서 목록", style="SubTitle.TLabel").pack(pady=(5, 5))

        # Transaction tree
        columns = ("거래 ID", "학생", "도서", "대출일", "반납 예정일", "상태")
        self.transaction_tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=10)
        for col in columns:
            self.transaction_tree.heading(col, text=col)
            self.transaction_tree.column(col, width=120, anchor="center")

        # Scrollbar for transaction tree
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.transaction_tree.yview)
        self.transaction_tree.configure(yscrollcommand=scrollbar.set)

        # Pack transaction tree and scrollbar
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)
        self.transaction_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.refresh_data()

    def refresh_data(self):
        """DB에서 데이터를 새로고침하고 드롭다운과 거래 목록을 업데이트합니다."""
        conn = get_db_connection()
        cursor = conn.cursor()

        # 데이터 딕셔너리로 미리 로드
        books = {row[0]: {'title': row[1], 'available': row[2]} for row in cursor.execute("SELECT book_id, title, available FROM books").fetchall()}
        students = {row[0]: {'name': row[1]} for row in cursor.execute("SELECT student_id, name FROM students").fetchall()}

        # 1. 학생 검색 업데이트
        student_values = [f"{sid} - {student['name']}" for sid, student in students.items()]
        self.student_search.set_values(student_values)

        # 2. 도서 검색 업데이트 (재고 있는 도서만)
        available_books_values = [f"{bid} - {book['title']} (재고: {book['available']})"
                                  for bid, book in books.items() if book['available'] > 0]
        self.book_search.set_values(available_books_values)

        # 3. 거래 목록 및 검색 업데이트 (대출 중인 거래만)
        self.transaction_tree.delete(*self.transaction_tree.get_children())
        issued_transactions_values = []

        transactions = cursor.execute("SELECT transaction_id, student_id, book_id, issue_date, due_date, status FROM transactions WHERE status = '대출'").fetchall()

        for row in transactions:
            tid, sid, bid, issue_date, due_date, status = row
            student_name = students.get(sid, {}).get('name', '알 수 없음')
            book_title = books.get(bid, {}).get('title', '알 수 없음')

            # Treeview insert
            self.transaction_tree.insert("", "end", values=(tid, student_name, book_title, issue_date, due_date, status))

            # Transaction search dropdown
            issued_transactions_values.append(f"{tid} - {student_name} - {book_title}")

        self.transaction_search.set_values(issued_transactions_values)
        conn.close()

    def issue_book(self):
        """도서 대출 처리 (DB 작업 포함)"""
        student_text = self.student_id_var.get().strip()
        book_text = self.book_id_var.get().strip()

        if not student_text or not book_text:
            show_error("오류", "학생 및 도서를 모두 선택해주세요.", self.parent)
            return

        try:
            student_id = student_text.split(" - ")[0] if " - " in student_text else student_text
            book_id = book_text.split(" - ")[0] if " - " in book_text else book_text
        except:
            show_error("오류", "유효한 학생 또는 도서 형식을 선택/입력해주세요.", self.parent)
            return

        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. 학생 ID 유효성 검사
        if not cursor.execute("SELECT 1 FROM students WHERE student_id = ?", (student_id,)).fetchone():
            show_error("오류", f"유효하지 않은 학생 ID: {student_id}", self.parent)
            conn.close()
            return

        # 2. 도서 정보 및 재고 확인
        book_info = cursor.execute("SELECT copies, available, title FROM books WHERE book_id = ?", (book_id,)).fetchone()

        if not book_info:
            show_error("오류", f"유효하지 않은 도서 ID: {book_id}", self.parent)
            conn.close()
            return

        copies, available, book_title = book_info

        if available <= 0:
            show_error("오류", f"선택한 도서 ('{book_title}')의 재고가 없습니다.", self.parent)
            conn.close()
            return

        # 3. 거래 생성
        transaction_id = generate_id("TXN", "transactions")
        issue_date = datetime.today().strftime("%Y-%m-%d")
        due_date = (datetime.today() + timedelta(days=14)).strftime("%Y-%m-%d")

        cursor.execute("""
            INSERT INTO transactions (transaction_id, student_id, book_id, issue_date, due_date, status, fine)
            VALUES (?, ?, ?, ?, ?, ?, 0)
        """, (transaction_id, student_id, book_id, issue_date, due_date, "대출"))

        # 4. 도서 재고 업데이트
        cursor.execute("UPDATE books SET available = available - 1 WHERE book_id = ?", (book_id,))

        conn.commit()
        conn.close()

        # 5. 후처리
        self.refresh_data()
        self.student_id_var.set("")
        self.book_id_var.set("")
        show_info("대출 성공", f"도서 대출이 성공적으로 완료되었습니다!\n거래 ID: {transaction_id}", self.parent)

    def return_book(self):
        """도서 반납 처리 (DB 작업 포함)"""
        transaction_text = self.transaction_id_var.get().strip()

        if not transaction_text:
            show_error("오류", "반납할 거래를 선택하거나 입력해주세요.", self.parent)
            return

        try:
            transaction_id = transaction_text.split(" - ")[0] if " - " in transaction_text else transaction_text
        except:
            show_error("오류", "유효한 거래 ID를 선택하거나 입력해주세요.", self.parent)
            return

        conn = get_db_connection()
        cursor = conn.cursor()

        transaction = cursor.execute("SELECT book_id, due_date, status FROM transactions WHERE transaction_id = ?", (transaction_id,)).fetchone()

        if not transaction:
            show_error("오류", f"거래 ID ({transaction_id})를 찾을 수 없습니다.", self.parent)
            conn.close()
            return

        book_id, due_date_str, status = transaction

        if status != '대출':
            show_error("오류", "이미 반납되었거나 유효하지 않은 거래입니다.", self.parent)
            conn.close()
            return

        # 1. 연체료 계산
        due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
        today = datetime.today().date()
        fine = 0
        days_overdue = 0

        if today > due_date:
            days_overdue = (today - due_date).days
            fine = days_overdue * FINE_PER_DAY

        # 2. 거래 업데이트
        return_date = today.strftime("%Y-%m-%d")
        cursor.execute("""
            UPDATE transactions 
            SET status = '반납', return_date = ?, fine = ? 
            WHERE transaction_id = ?
        """, (return_date, fine, transaction_id))

        # 3. 도서 재고 업데이트
        cursor.execute("UPDATE books SET available = available + 1 WHERE book_id = ?", (book_id,))

        conn.commit()
        conn.close()

        # 4. 후처리
        self.refresh_data()
        self.transaction_id_var.set("")

        if fine > 0:
            show_info("반납 완료", f"도서가 성공적으로 반납되었습니다.\n**연체료: ₩{fine:,}** ({days_overdue}일 연체)", self.parent)
        else:
            show_info("반납 완료", "도서가 성공적으로 반납되었습니다.", self.parent)


class LibraryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("도서관 관리 시스템")
        self.geometry("1000x800")
        self.resizable(True, True)

        try:
            self.iconbitmap("library.ico")
        except:
            pass

        self.center_window()

        # SQLite DB 초기화 및 구조 업데이트 실행
        init_db()

        # Style configuration
        self.configure_styles()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.dashboard_tab = DashboardTab(self)
        self.book_tab = BookManagementTab(self)
        self.student_tab = StudentManagementTab(self)
        self.issue_return_tab = IssueReturnTab(self)

        self.notebook.add(self.dashboard_tab, text="대시보드")
        self.notebook.add(self.book_tab, text="도서 관리")
        self.notebook.add(self.student_tab, text="학생 관리")
        self.notebook.add(self.issue_return_tab, text="대출 / 반납")

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        # Bind close event
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 메뉴바 추가
        self.create_menubar()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def configure_styles(self):
        """기본 글꼴을 '맑은 고딕', 11pt로 설정하고 특정 요소 스타일을 정의합니다."""
        style = ttk.Style()
        style.theme_use('clam')

        # 1. 전역 기본 글꼴 설정 (모든 TWidget에 적용)
        default_font = ("맑은 고딕", 11)
        style.configure('.', font=default_font)

        # 2. Treeview Headings (헤더) 글꼴 설정
        style.configure("Treeview.Heading", font=("맑은 고딕", 11, "bold"))

        # 3. Notebook Tab 글꼴 설정
        style.configure('TNotebook.Tab', padding=[15, 8], font=default_font)

        # 4. 특정 제목/통계 스타일
        style.configure('Title.TLabel', font=("맑은 고딕", 14, "bold"))  # 메인 제목
        style.configure('SubTitle.TLabel', font=("맑은 고딕", 11, "bold"))  # 서브 제목 및 카드 타이틀
        style.configure('Stats.TLabel', font=("맑은 고딕", 18, "bold"))  # 대시보드 통계 숫자

    def create_menubar(self):
        """메뉴바를 생성하고 Excel Import/Export 기능을 추가합니다."""
        menubar = tk.Menu(self)

        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)

        # 데이터 관리 메뉴 (Excel 기능)
        data_menu = tk.Menu(file_menu, tearoff=0)
        data_menu.add_command(label="모든 데이터 Excel로 내보내기", command=self.export_all_data_to_excel)
        data_menu.add_command(label="Excel에서 BOOKS 데이터 가져오기", command=lambda: self.import_data_from_excel("books"))
        data_menu.add_command(label="Excel에서 STUDENTS 데이터 가져오기", command=lambda: self.import_data_from_excel("students"))

        file_menu.add_cascade(label="데이터 관리 (Excel)", menu=data_menu)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.on_closing)

        menubar.add_cascade(label="파일", menu=file_menu)
        self.config(menu=menubar)

    def export_all_data_to_excel(self):
        """모든 테이블(books, students, transactions)의 데이터를 Excel 파일로 내보냅니다."""

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="모든 데이터 Excel로 내보내기"
        )

        if not filepath:
            return

        conn = get_db_connection()
        try:
            # 1. BOOKS 테이블
            df_books = pd.read_sql_query("SELECT * FROM books", conn)
            # 2. STUDENTS 테이블
            df_students = pd.read_sql_query("SELECT * FROM students", conn)
            # 3. TRANSACTIONS 테이블
            df_transactions = pd.read_sql_query("SELECT * FROM transactions", conn)

            # Excel Writer를 사용하여 여러 시트에 쓰기
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df_books.to_excel(writer, sheet_name='BOOKS', index=False)
                df_students.to_excel(writer, sheet_name='STUDENTS', index=False)
                df_transactions.to_excel(writer, sheet_name='TRANSACTIONS', index=False)

            show_info("내보내기 성공", f"모든 데이터가 '{filepath}' 파일에 성공적으로 저장되었습니다.", self)

        except pd.errors.EmptyDataError:
            show_error("오류", "데이터베이스에 내보낼 데이터가 없습니다.", self)
        except Exception as e:
            show_error("오류", f"Excel 내보내기 중 오류가 발생했습니다: {e}", self)
        finally:
            conn.close()

    def import_data_from_excel(self, table_name):
        """Excel 파일에서 지정된 테이블로 데이터를 가져옵니다."""

        # 사용자에게 파일 선택 요청
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title=f"Excel에서 {table_name.upper()} 데이터 가져오기"
        )

        if not filepath:
            return

        sheet_name = table_name.upper()

        try:
            # 1. Excel 파일 읽기
            # header=0는 첫 번째 행을 컬럼 이름으로 사용함을 의미
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=0, engine='openpyxl')

            # DB 컬럼 순서 및 이름 확인 (컬럼 불일치 시 오류 방지)
            if table_name == "books":
                required_cols = ['book_id', 'title', 'author', 'category', 'copies', 'available', 'added_date']
            elif table_name == "students":
                required_cols = ['student_id', 'name', 'major', 'phone', 'enroll_date']
            else:
                show_error("오류", f"지원하지 않는 테이블 이름: {table_name}", self)
                return

            if not all(col in df.columns for col in required_cols):
                show_error("오류", f"Excel 시트 '{sheet_name}'의 컬럼 이름이 데이터베이스 스키마와 일치하지 않습니다.\n"
                                 f"필수 컬럼: {', '.join(required_cols)}", self)
                return

            # 필요한 컬럼만 선택하고 순서 맞추기
            df = df[required_cols]

            # 2. DB 연결 및 데이터 삽입 (덮어쓰기/갱신 모드)
            conn = get_db_connection()
            # if_exists='replace'를 사용하여 기존 데이터를 모두 삭제하고 덮어쓰기
            confirm = ask_yesno("데이터 가져오기 확인",
                                f"경고: Excel 파일의 '{sheet_name}' 시트의 데이터로\n"
                                f"기존 DB의 {sheet_name} 테이블 데이터가 **모두 덮어쓰기**됩니다.\n"
                                f"계속하시겠습니까? (현재 {len(df)}개의 행)", self)

            if not confirm:
                conn.close()
                return

            # 데이터프레임의 모든 ID 필드가 문자열인지 확인 (DB 스키마와 일치)
            if table_name == "books":
                id_col = 'book_id'
                # copies와 available 필드가 정수인지 확인
                df['copies'] = pd.to_numeric(df['copies'], errors='coerce').astype('Int64')
                df['available'] = pd.to_numeric(df['available'], errors='coerce').astype('Int64')
                # 결측치 확인
                if df[['copies', 'available']].isnull().values.any():
                    show_error("오류", "copies나 available 컬럼에 유효하지 않은 정수 값이 포함되어 있습니다.", self)
                    conn.close()
                    return
            elif table_name == "students":
                id_col = 'student_id'

            df[id_col] = df[id_col].astype(str)

            # SQLite에 쓰기
            df.to_sql(table_name, conn, if_exists='replace', index=False)

            # BOOK 테이블을 덮어쓰는 경우, AVAILABLE = COPIES 가정을 유지하기 위해 추가 업데이트
            if table_name == "books":
                # 'available' 컬럼이 없는 경우를 대비하여 다시 확인
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT available FROM books LIMIT 1")
                except sqlite3.OperationalError:
                    cursor.execute("ALTER TABLE books ADD COLUMN available INTEGER NOT NULL DEFAULT 0")

                # copies와 available을 동일하게 설정하여 재고 오류 방지 (기존 대출 기록 무시)
                conn.execute("UPDATE books SET available = copies")

            conn.commit()
            conn.close()

            show_info("가져오기 성공", f"{sheet_name} 테이블에 {len(df)}개의 데이터가 성공적으로 가져와져 덮어쓰기 되었습니다.", self)
            self.on_tab_change(None)  # 모든 탭 새로고침

        except FileNotFoundError:
            show_error("오류", "파일을 찾을 수 없습니다.", self)
        except KeyError:
            show_error("오류", f"선택한 파일에 '{sheet_name}' 시트가 존재하지 않습니다.", self)
        except pd.errors.ParserError as e:
            show_error("오류", f"Excel 파일 파싱 중 오류 발생: {e}", self)
        except Exception as e:
            show_error("오류", f"데이터 가져오기 중 오류가 발생했습니다: {e}", self)
            if 'conn' in locals() and conn:
                conn.close()

    def on_tab_change(self, event):
        # NOTE: event가 None일 때 (import_data_from_excel에서 호출) 오류 방지를 위해 처리
        if event:
            current_tab = event.widget.tab(event.widget.index("current"))["text"]
        else:
            # 강제 새로고침 시 모든 탭 데이터 새로고침
            current_tab = None

        if current_tab == "대시보드" or current_tab is None:
            self.dashboard_tab.refresh()
        if current_tab == "대출 / 반납" or current_tab is None:
            self.issue_return_tab.refresh_data()
        if current_tab == "학생 관리" or current_tab is None:
            self.student_tab.refresh_students()
        if current_tab == "도서 관리" or current_tab is None:
            self.book_tab.refresh_books()

    def on_closing(self):
        """Handle window closing"""
        if ask_yesno("종료 확인", "정말로 프로그램을 종료하시겠습니까?", self):
            self.destroy()

    def save_all(self):
        # DB 방식에서는 save_all 대신 각 함수에서 commit을 수행합니다.
        pass


class DashboardTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent.notebook)
        self.parent = parent

        # Title (Style 적용)
        title_frame = ttk.Frame(self)
        title_frame.pack(fill="x", pady=10)
        ttk.Label(title_frame, text="📚 도서관 대시보드", style="Title.TLabel").pack()

        # Stats container
        stats_container = ttk.Frame(self)
        stats_container.pack(expand=True, fill="both", padx=50, pady=10)

        # Create stats cards in a grid
        self.create_stat_cards(stats_container)

        # Recent activity section
        activity_frame = ttk.LabelFrame(stats_container, text="최근 활동 기록", padding=10)
        activity_frame.pack(fill="x", pady=(5, 0))

        # Tkinter Text 위젯의 폰트 설정
        self.activity_text = tk.Text(activity_frame, height=10, wrap="word", state="disabled", font=("맑은 고딕", 11))
        activity_scrollbar = ttk.Scrollbar(activity_frame, orient="vertical", command=self.activity_text.yview)
        self.activity_text.configure(yscrollcommand=activity_scrollbar.set)

        self.activity_text.pack(side="left", fill="both", expand=True)
        activity_scrollbar.pack(side="right", fill="y")

        self.refresh()

    def create_stat_card(self, parent, title, value, color):
        """Create a statistics card (Style 적용)"""
        frame = ttk.LabelFrame(parent, text="", padding=5)

        title_label = ttk.Label(frame, text=title, style="SubTitle.TLabel")  # 11pt bold
        title_label.pack()

        value_label = ttk.Label(frame, text=value, style="Stats.TLabel")  # 18pt bold
        value_label.pack(pady=(1, 0))

        # Store reference to value label for updating
        frame.value_label = value_label

        return frame

    def create_stat_cards(self, parent):
        """Create statistics cards"""
        # First row
        row1 = ttk.Frame(parent)
        row1.pack(fill="x", pady=10)

        self.total_books_frame = self.create_stat_card(row1, "📚 총 도서 수", "0", "blue")
        self.total_books_frame.pack(side="left", fill="x", expand=True, padx=5)

        self.total_students_frame = self.create_stat_card(row1, "👥 총 학생 수", "0", "green")
        self.total_students_frame.pack(side="left", fill="x", expand=True, padx=5)

        # Second row
        row2 = ttk.Frame(parent)
        row2.pack(fill="x", pady=10)

        self.issued_books_frame = self.create_stat_card(row2, "📖 대출 중인 도서", "0", "orange")
        self.issued_books_frame.pack(side="left", fill="x", expand=True, padx=5)

        self.overdue_books_frame = self.create_stat_card(row2, "⚠️ 연체 도서", "0", "red")
        self.overdue_books_frame.pack(side="left", fill="x", expand=True, padx=5)

        # Third row
        row3 = ttk.Frame(parent)
        row3.pack(fill="x", pady=10)

        self.fine_collected_frame = self.create_stat_card(row3, "💰 총 누적 연체료", "0", "purple")
        self.fine_collected_frame.pack(side="left", fill="x", expand=True, padx=5)

        self.available_books_frame = self.create_stat_card(row3, "✅ 대출 가능 도서", "0", "teal")
        self.available_books_frame.pack(side="left", fill="x", expand=True, padx=5)

    def refresh(self):
        conn = get_db_connection()
        cursor = conn.cursor()

        # Load necessary data
        books_data = {row[0]: {'copies': row[1], 'available': row[2]} for row in cursor.execute("SELECT book_id, copies, available FROM books").fetchall()}
        students_data = cursor.execute("SELECT COUNT(*) FROM students").fetchone()[0]

        # Calculate stats
        total_books = sum(book["copies"] for book in books_data.values())
        total_students = students_data

        # Transactions query
        transactions_data = cursor.execute("SELECT due_date, status, fine FROM transactions").fetchall()

        issued_books = 0
        overdue_books = 0
        fine_collected = 0

        today = datetime.today().date()

        for row in transactions_data:
            due_date_str, status, fine = row
            fine_collected += fine if fine else 0

            if status == "대출":
                issued_books += 1
                try:
                    due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                    if today > due_date:
                        overdue_books += 1
                except (ValueError, TypeError):
                    # due_date 형식이 잘못된 경우 무시
                    pass

        available_books = sum(book["available"] for book in books_data.values())

        conn.close()

        # Update stat cards
        self.total_books_frame.value_label.config(text=str(total_books))
        self.total_students_frame.value_label.config(text=str(total_students))
        self.issued_books_frame.value_label.config(text=str(issued_books))
        self.overdue_books_frame.value_label.config(text=str(overdue_books))
        self.fine_collected_frame.value_label.config(text=f"₩{fine_collected:,}")
        self.available_books_frame.value_label.config(text=str(available_books))

        # Update recent activity
        self.update_recent_activity()

    def update_recent_activity(self):
        """Update recent activity section"""
        conn = get_db_connection()
        cursor = conn.cursor()

        # Get names for display
        students = {row[0]: row[1] for row in cursor.execute("SELECT student_id, name FROM students").fetchall()}
        books = {row[0]: row[1] for row in cursor.execute("SELECT book_id, title FROM books").fetchall()}

        # Get recent transactions (last 5, using PK for order)
        recent_transactions = cursor.execute(
            "SELECT transaction_id, student_id, book_id, issue_date, return_date, status FROM transactions ORDER BY transaction_id DESC LIMIT 5").fetchall()

        conn.close()

        self.activity_text.config(state="normal")
        self.activity_text.delete("1.0", "end")

        for row in recent_transactions:
            tid, sid, bid, issue_date, return_date, status = row
            student_name = students.get(sid, '알 수 없는 학생')
            book_title = books.get(bid, '알 수 없는 도서')

            if status == '대출':
                activity = f"📖 {student_name} 학생이 '{book_title}' 도서를 대출 - {issue_date}\n"
            else:
                activity = f"✅ {student_name} 학생이 '{book_title}' 도서를 반납 - {return_date if return_date else 'N/A'}\n"

            self.activity_text.insert("end", activity)

        if not recent_transactions:
            self.activity_text.insert("end", "최근 활동 기록이 없습니다.")

        self.activity_text.config(state="disabled")


class BookManagementTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent.notebook)
        self.parent = parent

        # Top frame
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", pady=5)

        self.search_var = tk.StringVar()
        ttk.Label(top_frame, text="검색:").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self.on_search)

        # --- 버튼 추가 ---
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side="right")

        ttk.Button(button_frame, text="신규 도서 추가", command=self.open_add_book_window).pack(side="left", padx=5)
        ttk.Button(button_frame, text="선택 도서 수정", command=self.edit_book).pack(side="left", padx=5)
        ttk.Button(button_frame, text="선택 도서 삭제", command=self.delete_book).pack(side="left", padx=5)
        ttk.Button(button_frame, text="새로고침", command=self.refresh_books).pack(side="left")
        # -----------------

        # Tree frame with scrollbar
        tree_frame = ttk.Frame(self)
        tree_frame.pack(expand=True, fill="both", pady=10)

        columns = ("ID", "도서명", "저자", "분류", "총 수량", "재고", "등록일")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        # Vertical scrollbar
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        # Horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        # Pack tree and scrollbars
        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        # Context menu
        self.menu = tk.Menu(self, tearoff=0, font=("맑은 고딕", 11))  # Tk Menu 폰트 개별 설정
        self.menu.add_command(label="도서 정보 수정", command=self.edit_book)
        self.menu.add_command(label="도서 삭제", command=self.delete_book)
        self.tree.bind("<Button-3>", self.show_menu)
        self.tree.bind("<Double-1>", lambda e: self.edit_book())  # 더블클릭으로 수정 호출

        self.refresh_books()

    def refresh_books(self):
        self.tree.delete(*self.tree.get_children())
        search_text = self.search_var.get().lower()

        conn = get_db_connection()
        cursor = conn.cursor()

        # DB에서 도서 목록 조회
        books_data = cursor.execute("SELECT * FROM books").fetchall()

        conn.close()

        for row in books_data:
            book_id, title, author, category, copies, available, added_date = row

            # 검색 필터링
            if (search_text in book_id.lower() or
                    search_text in title.lower() or
                    search_text in author.lower() or
                    (category and search_text in category.lower())):
                self.tree.insert("", "end", iid=book_id, values=(
                    book_id, title, author, category, copies, available, added_date
                ))

    def on_search(self, event):
        self.refresh_books()

    def show_menu(self, event):
        selected = self.tree.identify_row(event.y)
        if selected:
            self.tree.selection_set(selected)
            self.menu.post(event.x_root, event.y_root)

    def open_add_book_window(self):
        BookForm(self.parent, self, mode="add")

    def edit_book(self):
        selected = self.tree.selection()
        if not selected:
            show_warning("경고", "수정할 도서를 선택해주세요.", self.parent)
            return
        book_id = selected[0]
        BookForm(self.parent, self, mode="edit", book_id=book_id)

    def delete_book(self):
        selected = self.tree.selection()
        if not selected:
            show_warning("경고", "삭제할 도서를 선택해주세요.", self.parent)
            return
        book_id = selected[0]

        conn = get_db_connection()
        cursor = conn.cursor()

        # Check if book is currently issued
        issued_count = cursor.execute("SELECT COUNT(*) FROM transactions WHERE book_id = ? AND status = '대출'", (book_id,)).fetchone()[0]

        if issued_count > 0:
            show_error("오류", "현재 대출 중인 도서는 삭제할 수 없습니다.", self.parent)
            conn.close()
            return

        confirm = ask_yesno("삭제 확인", f"도서 ID {book_id}를 정말로 삭제하시겠습니까?", self.parent)
        if confirm:
            try:
                cursor.execute("DELETE FROM books WHERE book_id = ?", (book_id,))
                conn.commit()
                self.refresh_books()
                show_info("삭제 완료", "도서가 성공적으로 삭제되었습니다.", self.parent)
            except sqlite3.Error as e:
                show_error("DB 오류", f"도서 삭제 중 오류가 발생했습니다: {e}", self.parent)
                conn.rollback()
        conn.close()


class BookForm(tk.Toplevel):
    def __init__(self, master, parent_tab, mode="add", book_id=None):
        super().__init__(master)
        self.title("도서 정보 추가" if mode == "add" else "도서 정보 수정")
        self.transient(master)  # 부모 창 위에 유지
        self.grab_set()  # 다른 창 조작 방지
        self.parent_tab = parent_tab
        self.mode = mode
        self.book_id = book_id

        # 변수 정의
        self.title_var = tk.StringVar()
        self.author_var = tk.StringVar()
        self.category_var = tk.StringVar()
        self.copies_var = tk.StringVar(value="1")  # 기본값 1

        self.create_widgets()
        self.center_window()

        if self.mode == "edit" and self.book_id:
            self.load_book_data()

        # 창 닫기 이벤트 바인딩
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"+{x}+{y}")

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)

        # Labels and Entries
        fields = [
            ("도서명:", self.title_var, 0),
            ("저자:", self.author_var, 1),
            ("분류:", self.category_var, 2),
            ("총 수량:", self.copies_var, 3),
        ]

        for i, (label_text, var, row) in enumerate(fields):
            ttk.Label(main_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=5, padx=5)
            entry = ttk.Entry(main_frame, textvariable=var, width=40)
            entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)

        # 총 수량은 수정 모드에서 대출 중인 수량보다 작아질 수 없도록 제한해야 함
        if self.mode == "edit":
            ttk.Label(main_frame, text="(주의: 대출 중인 수량보다 작게 설정 불가)").grid(row=3, column=2, sticky="w", padx=5)

        # ID 표시 (수정 모드에서만)
        if self.mode == "edit":
            ttk.Label(main_frame, text="도서 ID:").grid(row=4, column=0, sticky="w", pady=5, padx=5)
            ttk.Label(main_frame, text=self.book_id).grid(row=4, column=1, sticky="w", pady=5, padx=5)

        # Button
        button_text = "도서 추가" if self.mode == "add" else "정보 수정"
        ttk.Button(main_frame, text=button_text, command=self.save_book).grid(row=5, column=0, columnspan=2, pady=15)

        main_frame.grid_columnconfigure(1, weight=1)

    def load_book_data(self):
        """수정 모드일 때 기존 도서 정보를 불러옵니다."""
        conn = get_db_connection()
        cursor = conn.cursor()
        book = cursor.execute("SELECT title, author, category, copies FROM books WHERE book_id = ?", (self.book_id,)).fetchone()
        conn.close()

        if book:
            self.title_var.set(book[0])
            self.author_var.set(book[1])
            self.category_var.set(book[2] if book[2] else "")
            self.copies_var.set(str(book[3]))
        else:
            show_error("오류", "도서 정보를 불러올 수 없습니다.", self)
            self.destroy()

    def save_book(self):
        """도서 정보를 저장하거나 업데이트합니다."""
        title = self.title_var.get().strip()
        author = self.author_var.get().strip()
        category = self.category_var.get().strip()
        copies_str = self.copies_var.get().strip()

        if not title or not author or not copies_str:
            show_error("오류", "도서명, 저자, 총 수량은 필수 입력 항목입니다.", self)
            return

        try:
            copies = int(copies_str)
            if copies < 0:
                raise ValueError
        except ValueError:
            show_error("오류", "총 수량은 유효한 양의 정수여야 합니다.", self)
            return

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            if self.mode == "add":
                # 신규 도서 추가
                book_id = generate_id("BK", "books")
                added_date = datetime.today().strftime("%Y-%m-%d")

                # 'copies'와 'available'을 동일하게 초기화
                cursor.execute("""
                    INSERT INTO books (book_id, title, author, category, copies, available, added_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (book_id, title, author, category, copies, copies, added_date))

                show_info("성공", f"도서 '{title}'이(가) 성공적으로 추가되었습니다. (ID: {book_id})", self)

            elif self.mode == "edit":
                # 도서 정보 수정
                # 1. 현재 대출 중인 도서 수 확인
                current_available, current_copies = cursor.execute(
                    "SELECT available, copies FROM books WHERE book_id = ?", (self.book_id,)
                ).fetchone()

                # 현재 대출 중인 수량 (issued = copies - available)
                issued_count = current_copies - current_available

                if copies < issued_count:
                    show_error("오류", f"새로운 총 수량 ({copies})은 현재 대출 중인 도서 수 ({issued_count})보다 작을 수 없습니다.", self)
                    return

                # 2. 새로운 'available' 재고 수량 계산
                # 새 재고 = 새 총 수량 - 대출 중인 수량
                new_available = copies - issued_count

                cursor.execute("""
                    UPDATE books SET title = ?, author = ?, category = ?, copies = ?, available = ?
                    WHERE book_id = ?
                """, (title, author, category, copies, new_available, self.book_id))

                show_info("성공", f"도서 '{title}' 정보가 성공적으로 수정되었습니다.", self)

            conn.commit()
            self.parent_tab.refresh_books()  # 부모 탭 새로고침
            self.destroy()

        except sqlite3.Error as e:
            show_error("DB 오류", f"DB 작업 중 오류가 발생했습니다: {e}", self)
            conn.rollback()
        finally:
            conn.close()

    def on_closing(self):
        """팝업 창 닫기 핸들러"""
        self.destroy()


class StudentManagementTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent.notebook)
        self.parent = parent

        # Top frame
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", pady=5)

        self.search_var = tk.StringVar()
        ttk.Label(top_frame, text="검색:").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self.on_search)

        # --- 버튼 추가 ---
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side="right")

        ttk.Button(button_frame, text="신규 학생 추가", command=self.open_add_student_window).pack(side="left", padx=5)
        ttk.Button(button_frame, text="선택 학생 수정", command=self.edit_student).pack(side="left", padx=5)
        ttk.Button(button_frame, text="선택 학생 삭제", command=self.delete_student).pack(side="left", padx=5)
        ttk.Button(button_frame, text="새로고침", command=self.refresh_students).pack(side="left")
        # -----------------

        # Tree frame with scrollbar
        tree_frame = ttk.Frame(self)
        tree_frame.pack(expand=True, fill="both", pady=10)

        columns = ("ID", "이름", "전공", "연락처", "등록일")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        # Pack tree and scrollbars
        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        # Context menu
        self.menu = tk.Menu(self, tearoff=0, font=("맑은 고딕", 11))
        self.menu.add_command(label="학생 정보 수정", command=self.edit_student)
        self.menu.add_command(label="학생 삭제", command=self.delete_student)
        self.tree.bind("<Button-3>", self.show_menu)
        self.tree.bind("<Double-1>", lambda e: self.edit_student())

        self.refresh_students()

    def refresh_students(self):
        self.tree.delete(*self.tree.get_children())
        search_text = self.search_var.get().lower()

        conn = get_db_connection()
        cursor = conn.cursor()

        # DB에서 학생 목록 조회
        students_data = cursor.execute("SELECT student_id, name, major, phone, enroll_date FROM students").fetchall()

        conn.close()

        for row in students_data:
            student_id, name, major, phone, enroll_date = row

            # 검색 필터링
            if (search_text in student_id.lower() or
                    search_text in name.lower() or
                    (major and search_text in major.lower()) or
                    (phone and search_text in phone.lower())):
                self.tree.insert("", "end", iid=student_id, values=(
                    student_id, name, major, phone, enroll_date
                ))

    def on_search(self, event):
        self.refresh_students()

    def show_menu(self, event):
        selected = self.tree.identify_row(event.y)
        if selected:
            self.tree.selection_set(selected)
            self.menu.post(event.x_root, event.y_root)

    def open_add_student_window(self):
        StudentForm(self.parent, self, mode="add")

    def edit_student(self):
        selected = self.tree.selection()
        if not selected:
            show_warning("경고", "수정할 학생을 선택해주세요.", self.parent)
            return
        student_id = selected[0]
        StudentForm(self.parent, self, mode="edit", student_id=student_id)

    def delete_student(self):
        selected = self.tree.selection()
        if not selected:
            show_warning("경고", "삭제할 학생을 선택해주세요.", self.parent)
            return
        student_id = selected[0]

        conn = get_db_connection()
        cursor = conn.cursor()

        # Check if student has outstanding issues
        issued_count = cursor.execute("SELECT COUNT(*) FROM transactions WHERE student_id = ? AND status = '대출'", (student_id,)).fetchone()[0]

        if issued_count > 0:
            show_error("오류", f"대출 중인 도서가 있는 학생은 삭제할 수 없습니다. (대출 건수: {issued_count})", self.parent)
            conn.close()
            return

        confirm = ask_yesno("삭제 확인", f"학생 ID {student_id}를 정말로 삭제하시겠습니까?", self.parent)
        if confirm:
            try:
                cursor.execute("DELETE FROM students WHERE student_id = ?", (student_id,))
                conn.commit()
                self.refresh_students()
                show_info("삭제 완료", "학생 정보가 성공적으로 삭제되었습니다.", self.parent)
            except sqlite3.Error as e:
                show_error("DB 오류", f"학생 삭제 중 오류가 발생했습니다: {e}", self.parent)
                conn.rollback()

        conn.close()


class StudentForm(tk.Toplevel):
    def __init__(self, master, parent_tab, mode="add", student_id=None):
        super().__init__(master)
        self.title("학생 정보 추가" if mode == "add" else "학생 정보 수정")
        self.transient(master)
        self.grab_set()
        self.parent_tab = parent_tab
        self.mode = mode
        self.student_id = student_id

        # 변수 정의
        self.name_var = tk.StringVar()
        self.major_var = tk.StringVar()
        self.phone_var = tk.StringVar()

        self.create_widgets()
        self.center_window()

        if self.mode == "edit" and self.student_id:
            self.load_student_data()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"+{x}+{y}")

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)

        # Labels and Entries
        fields = [
            ("이름:", self.name_var, 0),
            ("전공:", self.major_var, 1),
            ("연락처:", self.phone_var, 2),
        ]

        for i, (label_text, var, row) in enumerate(fields):
            ttk.Label(main_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=5, padx=5)
            entry = ttk.Entry(main_frame, textvariable=var, width=40)
            entry.grid(row=row, column=1, sticky="ew", pady=5, padx=5)

        # ID 표시 (수정 모드에서만)
        if self.mode == "edit":
            ttk.Label(main_frame, text="학생 ID:").grid(row=3, column=0, sticky="w", pady=5, padx=5)
            ttk.Label(main_frame, text=self.student_id).grid(row=3, column=1, sticky="w", pady=5, padx=5)

        # Button
        button_text = "학생 추가" if self.mode == "add" else "정보 수정"
        ttk.Button(main_frame, text=button_text, command=self.save_student).grid(row=4, column=0, columnspan=2, pady=15)

        main_frame.grid_columnconfigure(1, weight=1)

    def load_student_data(self):
        """수정 모드일 때 기존 학생 정보를 불러옵니다."""
        conn = get_db_connection()
        cursor = conn.cursor()
        student = cursor.execute("SELECT name, major, phone FROM students WHERE student_id = ?", (self.student_id,)).fetchone()
        conn.close()

        if student:
            self.name_var.set(student[0])
            self.major_var.set(student[1] if student[1] else "")
            self.phone_var.set(student[2] if student[2] else "")
        else:
            show_error("오류", "학생 정보를 불러올 수 없습니다.", self)
            self.destroy()

    def save_student(self):
        """학생 정보를 저장하거나 업데이트합니다."""
        name = self.name_var.get().strip()
        major = self.major_var.get().strip()
        phone = self.phone_var.get().strip()

        if not name:
            show_error("오류", "이름은 필수 입력 항목입니다.", self)
            return

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            if self.mode == "add":
                # 신규 학생 추가
                student_id = generate_id("STD", "students")
                enroll_date = datetime.today().strftime("%Y-%m-%d")

                cursor.execute("""
                    INSERT INTO students (student_id, name, major, phone, enroll_date)
                    VALUES (?, ?, ?, ?, ?)
                """, (student_id, name, major, phone, enroll_date))

                show_info("성공", f"학생 '{name}'이(가) 성공적으로 등록되었습니다. (ID: {student_id})", self)

            elif self.mode == "edit":
                # 학생 정보 수정
                cursor.execute("""
                    UPDATE students SET name = ?, major = ?, phone = ?
                    WHERE student_id = ?
                """, (name, major, phone, self.student_id))

                show_info("성공", f"학생 '{name}' 정보가 성공적으로 수정되었습니다.", self)

            conn.commit()
            self.parent_tab.refresh_students()
            self.destroy()

        except sqlite3.Error as e:
            show_error("DB 오류", f"DB 작업 중 오류가 발생했습니다: {e}", self)
            conn.rollback()
        finally:
            conn.close()

    def on_closing(self):
        """팝업 창 닫기 핸들러"""
        self.destroy()


# --- Main Execution ---

if __name__ == "__main__":
    try:
        app = LibraryApp()
        app.mainloop()
    except Exception as e:
        # Tkinter 오류 발생 시 콘솔에 출력
        print(f"An unexpected error occurred: {e}")
        # messagebox.showerror("치명적인 오류", f"프로그램 실행 중 오류가 발생했습니다: {e}")