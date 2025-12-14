"""
studentMarks.py 학생 성적 관리 Ver 1.1_251214
"""
import sqlite3
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = "Malgun Gothic"

DB_NAME = "studentssk.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS students
        (roll_no INTEGER PRIMARY KEY,
        name TEXT NOT NULL,
        class TEXT,
        marks_sub1 INTEGER DEFAULT 0,
        marks_sub2 INTEGER DEFAULT 0,
        marks_sub3 INTEGER DEFAULT 0,
        marks_sub4 INTEGER DEFAULT 0,
        marks_sub5 INTEGER DEFAULT 0,
        marks_sub6 INTEGER DEFAULT 0,
        marks_sub7 INTEGER DEFAULT 0,
        marks_sub8 INTEGER DEFAULT 0,
        marks_sub9 INTEGER DEFAULT 0,
        marks_sub10 INTEGER DEFAULT 0,
        marks_sub11 INTEGER DEFAULT 0,
        marks_sub12 INTEGER DEFAULT 0,
        total INTEGER,
        percentage REAL,
        grades TEXT)"""
    )
    conn.commit()
    conn.close()

def insert_student(data):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO students(roll_no,name,class,marks_sub1,marks_sub2,marks_sub3,marks_sub4,marks_sub5,marks_sub6,
    marks_sub7,marks_sub8,marks_sub9,marks_sub10,marks_sub11,marks_sub12,total,percentage,grades)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, data)
    conn.commit()
    conn.close()

def update_student_db(data):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("""
    UPDATE students
    SET name=?, class=?, marks_sub1=?, marks_sub2=?, marks_sub3=?, marks_sub4=?, marks_sub5=?, marks_sub6=?, 
        marks_sub7=?, marks_sub8=?, marks_sub9=?, marks_sub10=?, marks_sub11=?, marks_sub12=?, total=?, percentage=?, grades=?
    WHERE roll_no=?
    """, data)
    conn.commit()
    conn.close()

def delete_student_db(roll_no):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("DELETE FROM students WHERE roll_no=?", (roll_no,))
    conn.commit()
    conn.close()

def fetch_student(roll_no):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT * FROM students WHERE roll_no=?", (roll_no,))
    row = cur.fetchone()
    conn.close()
    return row

def fetch_all_students():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT * FROM students ORDER BY roll_no")
    rows = cur.fetchall()
    conn.close()
    return rows

def calculate_result(marks):
    total = sum(marks)
    # 12 subjects, each out of 100 => max total=1200
    percentage = (total / 1200) * 100
    # grade logic
    if percentage >= 90:
        grade = "A+"
    elif percentage >= 80:
        grade = "A"
    elif percentage >= 70:
        grade = "B+"
    elif percentage >= 60:
        grade = "B"
    elif percentage >= 50:
        grade = "C"
    else:
        grade = "F"
    return total, round(percentage, 2), grade


# GUI Application
class StudentRMSapp:
    def __init__(self, root):
        self.root = root
        root.title("🧑‍💻 학생 성적 관리 시스템")
        root.geometry("1020x640")
        root.resizable(False, False)

        # --- GUI 스타일 설정 ---
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', font=('Malgun Gothic', 10))

        style.configure('Success.TButton', foreground='white', background='#4CAF50', font=('Malgun Gothic', 10, 'bold'))
        style.configure('Info.TButton', foreground='white', background='#2196F3', font=('Malgun Gothic', 10, 'bold'))
        style.configure('Danger.TButton', foreground='white', background='#F44336', font=('Malgun Gothic', 10, 'bold'))
        style.map('Success.TButton', background=[('active', '#388E3C')])
        style.map('Info.TButton', background=[('active', '#1976D2')])
        style.map('Danger.TButton', background=[('active', '#D32F2F')])

        style.configure('Form.TLabelframe', background='#f0f0f0',font=('Malgun Gothic', 11, 'bold'))
        style.configure('Form.TLabelframe.Label', background='#f0f0f0',font=('Malgun Gothic', 11, 'bold'))

        self.var_roll = tk.StringVar()
        self.var_name = tk.StringVar()
        self.var_class = tk.StringVar()
        self.var_m1 = tk.StringVar()
        self.var_m2 = tk.StringVar()
        self.var_m3 = tk.StringVar()
        self.var_m4 = tk.StringVar()
        self.var_m5 = tk.StringVar()
        self.var_m6 = tk.StringVar()
        self.var_m7 = tk.StringVar()
        self.var_m8 = tk.StringVar()
        self.var_m9 = tk.StringVar()
        self.var_m10 = tk.StringVar()
        self.var_m11 = tk.StringVar()
        self.var_m12 = tk.StringVar()

        # Top Frame
        form_frame = ttk.LabelFrame(root, text="📝 학생 정보 및 성적 입력", style='Form.TLabelframe')
        form_frame.place(x=10, y=10, width=500, height=370)

        ttk.Label(form_frame, text="✅ 학번 (Roll No):").grid(row=0, column=0, padx=6, pady=8, sticky="W")
        ttk.Entry(form_frame, textvariable=self.var_roll, width=15).grid(row=0, column=1, padx=6, pady=8)

        ttk.Label(form_frame, text="👤 이름 (Name):").grid(row=0, column=2, padx=6, pady=8, sticky="W")
        ttk.Entry(form_frame, textvariable=self.var_name, width=14).grid(row=0, column=3, padx=6, pady=8)

        ttk.Label(form_frame, text="🏫 학급 (Class):").grid(row=1, column=0, padx=6, pady=8, sticky="w")
        ttk.Entry(form_frame, textvariable=self.var_class, width=15).grid(row=1, column=1, padx=6, pady=8)

        marks_frame = ttk.LabelFrame(form_frame, text="점수 입력 (0-100점)")
        marks_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5)

        ttk.Label(marks_frame, text="과목 1:").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m1, width=10).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 2:").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m2, width=10).grid(row=0, column=3, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 3:").grid(row=0, column=4, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m3, width=10).grid(row=0, column=5, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 4:").grid(row=1, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m4, width=10).grid(row=1, column=1, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 5:").grid(row=1, column=2, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m5, width=10).grid(row=1, column=3, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 6:").grid(row=1, column=4, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m6, width=10).grid(row=1, column=5, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 7:").grid(row=2, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m7, width=10).grid(row=2, column=1, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 8:").grid(row=2, column=2, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m8, width=10).grid(row=2, column=3, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목 9:").grid(row=2, column=4, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m9, width=10).grid(row=2, column=5, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목10:").grid(row=3, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m10, width=10).grid(row=3, column=1, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목11:").grid(row=3, column=2, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m11, width=10).grid(row=3, column=3, padx=6, pady=6)
        ttk.Label(marks_frame, text="과목12:").grid(row=3, column=4, padx=6, pady=6, sticky="w")
        ttk.Entry(marks_frame, textvariable=self.var_m12, width=10).grid(row=3, column=5, padx=6, pady=6)

        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, pady=10)

        ttk.Button(btn_frame, text="➕ 학생 추가", style='Success.TButton', width=15, command=self.add_student).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(btn_frame, text="✏️ 학생 수정", style='Info.TButton', width=15, command=self.update_student).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(btn_frame, text="❌ 학생 삭제", style='Danger.TButton', width=15, command=self.delete_student).grid(row=0, column=2, padx=4, pady=4)

        ttk.Button(btn_frame, text="🔍 학번 검색", width=15, command=self.search_student).grid(row=1, column=0, padx=4, pady=4)
        ttk.Button(btn_frame, text="🧹 필드 초기화", width=15, command=self.clear_fields).grid(row=1, column=1, padx=4, pady=4)

        control_frame = ttk.Labelframe(root, text="📊 결과 보기 및 데이터 관리",style='Form.TLabelframe')
        control_frame.place(x=540, y=10, width=470, height=370)

        ttk.Button(control_frame, text="📋 전체 학생 보기", command=self.view_all, width=20).pack(side="top", pady=(40, 10), ipadx=5, ipady=5)
        ttk.Button(control_frame, text="📤 CSV로 내보내기", command=self.export_csv, width=20).pack(side="top", pady=10, ipadx=5, ipady=5)
        ttk.Button(control_frame, text="📈 성적 분포 그래프", command=self.plot_grades, width=20).pack(side="top", pady=10, ipadx=5, ipady=5)
        ttk.Button(control_frame, text="🚪 종료", command=root.quit, width=20).pack(side="top", pady=10, ipadx=5, ipady=5)

        table_frame = ttk.Frame(root)
        table_frame.place(x=10, y=390, width=1000, height=230)

        columns = ("roll_no", "name", "class", "m1", "m2", "m3", "m4", "m5", "m6",
                   "m7", "m8", "m9", "m10", "m11", "m12", "total", "percentage", "grade")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=9)

        headings_map = {
            "roll_no": "학번", "name": "이름", "class": "반",
            "m1": "과목1", "m2": "과목2", "m3": "과목3", "m4": "과목4", "m5": "과목5", "m6": "과목6",
            "m7": "과목7", "m8": "과목8", "m9": "과목9", "m10": "과목10", "m11": "과목11", "m12": "과목12",
            "total": "총점", "percentage": "평균(%)", "grade": "성적"
        }

        for col in columns:
            self.tree.heading(col, text=headings_map.get(col, col.title().replace("_", " ")))
            if col == "name":
                w = 90
            elif col == "percentage":
                w = 70
            elif col == "total":
                w = 70
            elif col in ["m1", "m2", "m3", "m4", "m5", "m6", "m7", "m8", "m9", "m10", "m11", "m12"]:
                w = 60
            else:
                w = 70
            self.tree.column(col, width=w, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.view_all()
        self.tree.bind("<Double-1>", self.on_treeview_double_click)

    def clear_fields(self):
        self.var_roll.set("")
        self.var_name.set("")
        self.var_class.set("")
        self.var_m1.set("")
        self.var_m2.set("")
        self.var_m3.set("")
        self.var_m4.set("")
        self.var_m5.set("")
        self.var_m6.set("")
        self.var_m7.set("")
        self.var_m8.set("")
        self.var_m9.set("")
        self.var_m10.set("")
        self.var_m11.set("")
        self.var_m12.set("")

    def validate_marks(self, *marks_str):
        marks = []
        for s in marks_str:
            s = s.strip()
            if s == "":
                val = 0
            else:
                try:
                    val = int(s)
                except ValueError:
                    raise ValueError("점수는 정수여야 합니다.")
                if not (0 <= val <= 100):
                    raise ValueError("점수는 0점에서 100점 사이여야 합니다.")
            marks.append(val)
        return marks

    def add_student(self):
        roll = self.var_roll.get().strip()
        name = self.var_name.get().strip()
        cls = self.var_class.get().strip()
        if roll == "" or name == "":
            messagebox.showerror("입력 오류", "학번과 이름은 필수 입력 항목입니다.")
            return
        try:
            roll_no = int(roll)
        except ValueError:
            messagebox.showerror("입력 오류", "학번은 정수여야 합니다.")
            return

        try:
            # 버그 수정: 12개 과목 모두 검증
            marks = self.validate_marks(
                self.var_m1.get(), self.var_m2.get(), self.var_m3.get(),
                self.var_m4.get(), self.var_m5.get(), self.var_m6.get(),
                self.var_m7.get(), self.var_m8.get(), self.var_m9.get(),
                self.var_m10.get(), self.var_m11.get(), self.var_m12.get()
            )
        except ValueError as e:
            messagebox.showerror("점수 오류", str(e))
            return

        total, percentage, grade = calculate_result(marks)

        if fetch_student(roll_no):
            messagebox.showerror("중복 오류", f"학번 {roll_no} 학생이 이미 존재합니다.")
            return

        data = (roll_no, name, cls, marks[0], marks[1], marks[2], marks[3], marks[4], marks[5],
                marks[6], marks[7], marks[8], marks[9], marks[10], marks[11], total, percentage, grade)
        try:
            insert_student(data)
            messagebox.showinfo("성공", "학생 정보가 성공적으로 추가되었습니다.")
            self.view_all()
            self.clear_fields()
        except Exception as e:
            messagebox.showerror("DB 오류", str(e))

    def update_student(self):
        roll = self.var_roll.get().strip()
        if roll == "":
            messagebox.showerror("입력 오류", "수정할 학생의 학번을 입력해 주세요.")
            return
        try:
            roll_no = int(roll)
        except ValueError:
            messagebox.showerror("입력 오류", "학번은 정수여야 합니다.")
            return
        if not fetch_student(roll_no):
            messagebox.showerror("찾을 수 없음", f"학번 {roll_no}인 학생을 찾을 수 없습니다.")
            return

        name = self.var_name.get().strip()
        cls = self.var_class.get().strip()
        if name == "":
            messagebox.showerror("입력 오류", "이름은 필수 입력 항목입니다.")
            return
        try:
            marks = self.validate_marks(
                self.var_m1.get(), self.var_m2.get(), self.var_m3.get(),
                self.var_m4.get(), self.var_m5.get(), self.var_m6.get(),
                self.var_m7.get(), self.var_m8.get(), self.var_m9.get(),
                self.var_m10.get(), self.var_m11.get(), self.var_m12.get()
            )
        except ValueError as e:
            messagebox.showerror("점수 오류", str(e))
            return

        total, percentage, grade = calculate_result(marks)
        # 버그 수정: 12개 과목 모두 포함
        data = (name, cls, marks[0], marks[1], marks[2], marks[3], marks[4], marks[5],
                marks[6], marks[7], marks[8], marks[9], marks[10], marks[11], total, percentage, grade, roll_no)
        try:
            update_student_db(data)
            messagebox.showinfo("성공", "학생 정보가 성공적으로 수정되었습니다.")
            self.view_all()
            self.clear_fields()
        except Exception as e:
            messagebox.showerror("DB 오류", str(e))

    def delete_student(self):
        roll = self.var_roll.get().strip()
        if roll == "":
            messagebox.showerror("입력 오류", "삭제할 학생의 학번을 입력해 주세요.")
            return
        try:
            roll_no = int(roll)
        except ValueError:
            messagebox.showerror("입력 오류", "학번은 정수여야 합니다.")
            return
        if not fetch_student(roll_no):
            messagebox.showerror("찾을 수 없음", f"학번 {roll_no}인 학생을 찾을 수 없습니다.")
            return

        confirm = messagebox.askyesno("삭제 확인", f"학번 {roll_no}인 학생을 정말로 삭제하시겠습니까?")
        if not confirm:
            return

        try:
            delete_student_db(roll_no)
            messagebox.showinfo("성공", "학생 정보가 성공적으로 삭제되었습니다.")
            self.view_all()
            self.clear_fields()
        except Exception as e:
            messagebox.showerror("DB 오류", str(e))

    def search_student(self):
        roll = self.var_roll.get().strip()
        if roll == "":
            messagebox.showerror("입력 오류", "검색할 학번을 입력해 주세요.")
            return
        try:
            roll_no = int(roll)
        except ValueError:
            messagebox.showerror("입력 오류", "학번은 정수여야 합니다.")
            return

        row = fetch_student(roll_no)
        if not row:
            messagebox.showinfo("찾을 수 없음", f"학번 {roll_no}인 학생을 찾을 수 없습니다.")
            return

        self.var_name.set(row[1])
        self.var_class.set(row[2])
        self.var_m1.set(str(row[3]))
        self.var_m2.set(str(row[4]))
        self.var_m3.set(str(row[5]))
        self.var_m4.set(str(row[6]))
        self.var_m5.set(str(row[7]))
        self.var_m6.set(str(row[8]))
        self.var_m7.set(str(row[9]))
        self.var_m8.set(str(row[10]))
        self.var_m9.set(str(row[11]))
        self.var_m10.set(str(row[12]))
        self.var_m11.set(str(row[13]))
        self.var_m12.set(str(row[14]))

        self.view_all()
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if int(vals[0]) == roll_no:
                self.tree.selection_set(item)
                self.tree.focus(item)
                self.tree.see(item)
                break

    def view_all(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = fetch_all_students()
        for r in rows:
            self.tree.insert("", "end", values=r)

    def on_treeview_double_click(self, event):
        item = self.tree.focus()
        if not item:
            return
        vals = self.tree.item(item, "values")
        self.var_roll.set(vals[0])
        self.var_name.set(vals[1])
        self.var_class.set(vals[2])
        self.var_m1.set(vals[3])
        self.var_m2.set(vals[4])
        self.var_m3.set(vals[5])
        self.var_m4.set(vals[6])
        self.var_m5.set(vals[7])
        self.var_m6.set(vals[8])
        self.var_m7.set(vals[9])
        self.var_m8.set(vals[10])
        self.var_m9.set(vals[11])
        self.var_m10.set(vals[12])
        self.var_m11.set(vals[13])
        self.var_m12.set(vals[14])

    def export_csv(self):
        rows = fetch_all_students()
        if not rows:
            messagebox.showinfo("데이터 없음", "내보낼 레코드가 없습니다.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
            title="CSV 파일로 저장"
        )
        if not file_path:
            return
        try:
            with open(file_path, mode="w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                header = ["학번", "이름", "반", "과목1", "과목2", "과목3", "과목4", "과목5", "과목6",
                          "과목7", "과목8", "과목9", "과목10", "과목11", "과목12", "총점", "평균(%)", "성적"]
                writer.writerow(header)
                for r in rows:
                    writer.writerow(r)
            messagebox.showinfo("내보내기 성공", f"레코드가 {file_path}에 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def plot_grades(self):
        rows = fetch_all_students()
        if not rows:
            messagebox.showinfo("데이터 없음", "그래프를 그릴 학생이 없습니다.")
            return

        names = [r[1] for r in rows]
        percentages = [r[16] for r in rows]

        plt.figure(figsize=(10, 5))
        plt.bar(names, percentages, color='#4CAF50')

        plt.rcParams['font.family'] = 'Malgun Gothic'
        plt.rcParams['axes.unicode_minus'] = False

        plt.title("학생별 성적 평균 분포")
        plt.xlabel("학생 이름")
        plt.ylabel("평균 (%)")

        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.show()


if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = StudentRMSapp(root)
    root.mainloop()