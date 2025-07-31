"""
student_score.py 학생 (성적) 관리 시스템   Ver 1.0_2507730
"""
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import sqlite3
import datetime

class StudentManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("학생 관리 시스템")
        self.root.geometry("600x700")
        self.root.configure(bg="#003366") # 메인 창 배경색

        # 데이터베이스 초기화 (테이블 및 뷰 생성)
        self._initialize_db()
        # 데이터가 없을 경우 샘플 데이터 자동 생성
        self._generate_sample_data()

        # 메인 타이틀
        tk.Label(self.root, text="학생 관리 시스템", bg="#003366", fg="white", font=("D2Coding", 24, "bold")).pack(pady=30)

        # 메인 버튼들
        button_font = ("D2Coding", 16, "bold")
        button_width = 25
        button_pady = 15

        tk.Button(self.root, text="학생 추가", command=self.add_student, bg="#008080", fg="white",
                  font=button_font, width=button_width, relief="raised", bd=3).pack(pady=button_pady)
        tk.Button(self.root, text="강좌 추가", command=self.add_course, bg="#008080", fg="white",
                  font=button_font, width=button_width, relief="raised", bd=3).pack(pady=button_pady)
        tk.Button(self.root, text="성적 할당", command=self.assign_grade, bg="#008080", fg="white",
                  font=button_font, width=button_width, relief="raised", bd=3).pack(pady=button_pady)

        tk.Button(self.root, text="업데이트", command=self.open_update_interface, bg="#FFD700", fg="black",
                  font=button_font, width=button_width, relief="raised", bd=3).pack(pady=button_pady)

        tk.Button(self.root, text="보기", command=self.open_view_interface, bg="#0074D9", fg="white",
                  font=button_font, width=button_width, relief="raised", bd=3).pack(pady=button_pady)

        tk.Button(self.root, text="종료", command=self.exit_program, bg="#FF5733", fg="white",
                  font=button_font, width=button_width, relief="raised", bd=3).pack(pady=button_pady + 20)

    # 데이터베이스 연결 함수
    def connect_to_db(self):
        return sqlite3.connect('student_ms.db')

    # 데이터베이스 초기화 (테이블 및 뷰 생성)
    def _initialize_db(self):
        conn = self.connect_to_db()
        cursor = conn.cursor()

        # Students 테이블 생성
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Students (
                student_id TEXT PRIMARY KEY UNIQUE,
                name TEXT NOT NULL,
                email TEXT NOT NULL
            )
        """)

        # Courses 테이블 생성
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Courses (
                course_id TEXT PRIMARY KEY UNIQUE,
                course_name TEXT NOT NULL,
                credits INTEGER NOT NULL
            )
        """)

        # Grades 테이블 생성
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Grades (
                grade_id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id TEXT,
                course_id TEXT,
                grade TEXT,
                FOREIGN KEY (student_id) REFERENCES Students(student_id) ON DELETE CASCADE,
                FOREIGN KEY (course_id) REFERENCES Courses(course_id) ON DELETE CASCADE
            )
        """)

        # Views (SQLite는 CREATE OR REPLACE VIEW를 지원)
        cursor.execute("""
            CREATE VIEW IF NOT EXISTS StudentCourseGrades AS
            SELECT
                s.student_id,
                s.name,
                s.email,
                c.course_name,
                g.grade
            FROM Students s
            LEFT JOIN Grades g ON s.student_id = g.student_id
            LEFT JOIN Courses c ON g.course_id = c.course_id;
        """)

        cursor.execute("""
            CREATE VIEW IF NOT EXISTS CourseStudents AS
            SELECT
                c.course_id,
                c.course_name,
                s.student_id,
                s.name,
                g.grade
            FROM Courses c
            LEFT JOIN Grades g ON c.course_id = g.course_id
            LEFT JOIN Students s ON g.student_id = s.student_id;
        """)

        conn.commit()
        conn.close()

    def _generate_sample_data(self):
        conn = self.connect_to_db()
        cursor = conn.cursor()

        # Students 테이블이 비어 있는지 확인
        cursor.execute("SELECT COUNT(*) FROM Students")
        if cursor.fetchone()[0] == 0:
            try:
                # 샘플 학생 데이터 삽입
                students_data = [
                    ("S001", "김철수", "kim.cs@example.com"),
                    ("S002", "이영희", "lee.yh@example.com"),
                    ("S003", "박민수", "park.ms@example.com")
                ]
                cursor.executemany("INSERT INTO Students (student_id, name, email) VALUES (?, ?, ?)", students_data)

                # 샘플 강좌 데이터 삽입
                courses_data = [
                    ("C101", "파이썬 프로그래밍", 3),
                    ("C102", "데이터베이스 개론", 3),
                    ("C103", "자료구조", 3)
                ]
                cursor.executemany("INSERT INTO Courses (course_id, course_name, credits) VALUES (?, ?, ?)", courses_data)

                # 샘플 성적 데이터 할당
                grades_data = [
                    ("S001", "C101", "A+"),
                    ("S001", "C102", "B"),
                    ("S002", "C101", "A"),
                    ("S002", "C103", "C+"),
                    ("S003", "C102", "B+"),
                    ("S003", "C103", "A-")
                ]
                cursor.executemany("INSERT INTO Grades (student_id, course_id, grade) VALUES (?, ?, ?)", grades_data)

                conn.commit()
                print("샘플 데이터가 성공적으로 생성되었습니다.")
            except sqlite3.IntegrityError as e:
                print(f"샘플 데이터가 이미 존재하거나 무결성 오류가 발생했습니다: {e}")
                conn.rollback() # 데이터가 이미 존재하거나 다른 무결성 문제가 발생하면 롤백
            except Exception as e:
                print(f"샘플 데이터 생성 중 오류 발생: {e}")
                conn.rollback()
        else:
            print("데이터베이스에 이미 데이터가 있습니다. 샘플 데이터 생성은 건너뜁니다.")
        conn.close()

    # 창 중앙 정렬 헬퍼 함수
    def _center_window(self, window, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    # CRUD 함수
    def add_student_to_db(self, student_id, name, email):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO Students (student_id, name, email) VALUES (?, ?, ?)", (student_id, name, email))
            conn.commit()
            messagebox.showinfo("성공", "학생이 성공적으로 추가되었습니다!", parent=self.root)
        except sqlite3.IntegrityError:
            messagebox.showerror("오류", "학생 ID가 이미 존재합니다!", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"학생 추가 중 오류 발생: {e}", parent=self.root)
        finally:
            conn.close()

    def add_course_to_db(self, course_id, course_name, credits):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO Courses (course_id, course_name, credits) VALUES (?, ?, ?)", (course_id, course_name, credits))
            conn.commit()
            messagebox.showinfo("성공", "강좌가 성공적으로 추가되었습니다!", parent=self.root)
        except sqlite3.IntegrityError:
            messagebox.showerror("오류", "강좌 ID가 이미 존재합니다!", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"강좌 추가 중 오류 발생: {e}", parent=self.root)
        finally:
            conn.close()

    def assign_grade_to_db(self, student_id, course_id, grade):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        try:
            # 먼저 학생과 강좌가 존재하는지 확인
            cursor.execute("SELECT 1 FROM Students WHERE student_id = ?", (student_id,))
            student_exists = cursor.fetchone()
            cursor.execute("SELECT 1 FROM Courses WHERE course_id = ?", (course_id,))
            course_exists = cursor.fetchone()

            if not student_exists:
                messagebox.showerror("오류", "존재하지 않는 학생 ID입니다.", parent=self.root)
                return
            if not course_exists:
                messagebox.showerror("오류", "존재하지 않는 강좌 ID입니다.", parent=self.root)
                return

            # 이미 성적이 할당되었는지 확인 (업데이트 대신)
            cursor.execute("SELECT 1 FROM Grades WHERE student_id = ? AND course_id = ?", (student_id, course_id))
            grade_exists = cursor.fetchone()

            if grade_exists:
                messagebox.showwarning("경고", "이 학생에게 이미 해당 강좌의 성적이 할당되어 있습니다. 업데이트 기능을 사용하세요.", parent=self.root)
            else:
                cursor.execute("INSERT INTO Grades (student_id, course_id, grade) VALUES (?, ?, ?)", (student_id, course_id, grade))
                conn.commit()
                messagebox.showinfo("성공", "성적이 성공적으로 할당되었습니다!", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"성적 할당 중 오류 발생: {e}", parent=self.root)
        finally:
            conn.close()

    def search_student_by_id(self, student_id):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        cursor.execute("SELECT student_id, name, email, course_name, grade FROM StudentCourseGrades WHERE student_id = ?", (student_id,))
        records = cursor.fetchall()
        conn.close()
        return records

    def search_course_by_id(self, course_id):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        cursor.execute("SELECT course_id, course_name, student_id, name, grade FROM CourseStudents WHERE course_id = ?", (course_id,))
        records = cursor.fetchall()
        conn.close()
        return records

    # 모든 학생 데이터 조회
    def get_all_students_details(self):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        cursor.execute("SELECT student_id, name, email, course_name, grade FROM StudentCourseGrades ORDER BY student_id, course_name")
        records = cursor.fetchall()
        conn.close()
        return records

    # 모든 강좌 데이터 조회
    def get_all_courses_details(self):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        cursor.execute("SELECT course_id, course_name, student_id, name, grade FROM CourseStudents ORDER BY course_id, name")
        records = cursor.fetchall()
        conn.close()
        return records

    # GUI 함수
    def add_student(self):
        def save_student():
            student_id = student_id_entry.get().strip()
            name = name_entry.get().strip()
            email = email_entry.get().strip()
            if student_id and name and email:
                self.add_student_to_db(student_id, name, email)
                student_window.destroy()
            else:
                messagebox.showerror("오류", "모든 필드는 필수입니다!", parent=student_window)

        student_window = tk.Toplevel(self.root)
        student_window.title("학생 추가")
        self._center_window(student_window, 500, 350)
        student_window.configure(bg="#D3D3D3")

        form_frame = tk.Frame(student_window, bg="#D3D3D3")
        form_frame.pack(pady=20)

        label_font = ("D2Coding", 14)
        entry_font = ("D2Coding", 14)
        entry_width = 30

        tk.Label(form_frame, text="학생 ID:", bg="#D3D3D3", font=label_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        student_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        student_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="이름:", bg="#D3D3D3", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        name_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        name_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="이메일:", bg="#D3D3D3", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        email_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        email_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        tk.Button(student_window, text="저장", command=save_student, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), width=15, relief="raised", bd=3).pack(pady=20)

    def add_course(self):
        def save_course():
            course_id = course_id_entry.get().strip()
            course_name = course_name_entry.get().strip()
            credits = credits_entry.get().strip()
            if course_id and course_name and credits.isdigit():
                self.add_course_to_db(course_id, course_name, int(credits))
                course_window.destroy()
            else:
                messagebox.showerror("오류", "모든 필드는 필수이며 학점은 숫자여야 합니다!", parent=course_window)

        course_window = tk.Toplevel(self.root)
        course_window.title("강좌 추가")
        self._center_window(course_window, 500, 350)
        course_window.configure(bg="#D3D3D3")

        form_frame = tk.Frame(course_window, bg="#D3D3D3")
        form_frame.pack(pady=20)

        label_font = ("D2Coding", 14)
        entry_font = ("D2Coding", 14)
        entry_width = 30

        tk.Label(form_frame, text="강좌 ID:", bg="#D3D3D3", font=label_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        course_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        course_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="강좌 이름:", bg="#D3D3D3", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        course_name_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        course_name_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="학점:", bg="#D3D3D3", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        credits_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        credits_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        tk.Button(course_window, text="저장", command=save_course, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), width=15, relief="raised", bd=3).pack(pady=20)

    def assign_grade(self):
        def save_grade():
            student_id = student_id_entry.get().strip()
            course_id = course_id_entry.get().strip()
            grade = grade_entry.get().strip()
            if student_id and course_id and grade:
                self.assign_grade_to_db(student_id, course_id, grade)
                grade_window.destroy()
            else:
                messagebox.showerror("오류", "모든 필드는 필수입니다!", parent=grade_window)

        grade_window = tk.Toplevel(self.root)
        grade_window.title("성적 할당")
        self._center_window(grade_window, 500, 350)
        grade_window.configure(bg="#D3D3D3")

        form_frame = tk.Frame(grade_window, bg="#D3D3D3")
        form_frame.pack(pady=20)

        label_font = ("D2Coding", 14)
        entry_font = ("D2Coding", 14)
        entry_width = 30

        tk.Label(form_frame, text="학생 ID:", bg="#D3D3D3", font=label_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        student_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        student_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="강좌 ID:", bg="#D3D3D3", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        course_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        course_id_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="성적:", bg="#D3D3D3", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        grade_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        grade_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        tk.Button(grade_window, text="저장", command=save_grade, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), width=15, relief="raised", bd=3).pack(pady=20)

    def update_student_in_db(self, student_id, name, email):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE Students SET name = ?, email = ? WHERE student_id = ?", (name, email, student_id))
            if cursor.rowcount == 0:
                messagebox.showwarning("경고", "업데이트할 학생 ID를 찾을 수 없습니다.", parent=self.root)
            else:
                conn.commit()
                messagebox.showinfo("성공", "학생 정보가 성공적으로 업데이트되었습니다!", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"학생 정보 업데이트 중 오류 발생: {e}", parent=self.root)
        finally:
            conn.close()

    def update_course_in_db(self, course_id, course_name, credits):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE Courses SET course_name = ?, credits = ? WHERE course_id = ?", (course_name, credits, course_id))
            if cursor.rowcount == 0:
                messagebox.showwarning("경고", "업데이트할 강좌 ID를 찾을 수 없습니다.", parent=self.root)
            else:
                conn.commit()
                messagebox.showinfo("성공", "강좌 정보가 성공적으로 업데이트되었습니다!", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"강좌 정보 업데이트 중 오류 발생: {e}", parent=self.root)
        finally:
            conn.close()

    def update_grade_in_db(self, student_id, course_id, grade):
        conn = self.connect_to_db()
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE Grades SET grade = ? WHERE student_id = ? AND course_id = ?", (grade, student_id, course_id))
            if cursor.rowcount == 0:
                messagebox.showwarning("경고", "업데이트할 성적 기록을 찾을 수 없습니다. 학생 ID와 강좌 ID를 확인하세요.", parent=self.root)
            else:
                conn.commit()
                messagebox.showinfo("성공", "성적이 성공적으로 업데이트되었습니다!", parent=self.root)
        except Exception as e:
            messagebox.showerror("오류", f"성적 업데이트 중 오류 발생: {e}", parent=self.root)
        finally:
            conn.close()

    # GUI 함수
    def update_student(self):
        def save_update_student():
            student_id = student_id_entry.get().strip()
            name = name_entry.get().strip()
            email = email_entry.get().strip()
            if student_id and name and email:
                self.update_student_in_db(student_id, name, email)
                update_window.destroy()
            else:
                messagebox.showerror("오류", "모든 필드는 필수입니다!", parent=update_window)

        update_window = tk.Toplevel(self.root)
        update_window.title("학생 정보 업데이트")
        self._center_window(update_window, 500, 350)
        update_window.configure(bg="#D3D3D3")

        form_frame = tk.Frame(update_window, bg="#D3D3D3")
        form_frame.pack(pady=20)

        label_font = ("D2Coding", 14)
        entry_font = ("D2Coding", 14)
        entry_width = 30

        tk.Label(form_frame, text="학생 ID:", bg="#D3D3D3", font=label_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        student_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        student_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="이름:", bg="#D3D3D3", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        name_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        name_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="이메일:", bg="#D3D3D3", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        email_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        email_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        tk.Button(update_window, text="저장", command=save_update_student, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), width=15, relief="raised", bd=3).pack(pady=20)

    def update_course(self):
        def save_update_course():
            course_id = course_id_entry.get().strip()
            course_name = course_name_entry.get().strip()
            credits = credits_entry.get().strip()
            if course_id and course_name and credits.isdigit():
                self.update_course_in_db(course_id, course_name, int(credits))
                update_window.destroy()
            else:
                messagebox.showerror("오류", "모든 필드는 필수이며 학점은 숫자여야 합니다!", parent=update_window)

        update_window = tk.Toplevel(self.root)
        update_window.title("강좌 정보 업데이트")
        self._center_window(update_window, 500, 350)
        update_window.configure(bg="#D3D3D3")

        form_frame = tk.Frame(update_window, bg="#D3D3D3")
        form_frame.pack(pady=20)

        label_font = ("D2Coding", 14)
        entry_font = ("D2Coding", 14)
        entry_width = 30

        tk.Label(form_frame, text="강좌 ID:", bg="#D3D3D3", font=label_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        course_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        course_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="강좌 이름:", bg="#D3D3D3", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        course_name_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        course_name_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="학점:", bg="#D3D3D3", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        credits_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        credits_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        tk.Button(update_window, text="저장", command=save_update_course, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), width=15, relief="raised", bd=3).pack(pady=20)

    def update_grade(self):
        def save_update_grade():
            student_id = student_id_entry.get().strip()
            course_id = course_id_entry.get().strip()
            grade = grade_entry.get().strip()

            if student_id and course_id and grade:
                self.update_grade_in_db(student_id, course_id, grade)
                update_window.destroy()
            else:
                messagebox.showerror("오류", "모든 필드는 필수입니다!", parent=update_window)

        update_window = tk.Toplevel(self.root)
        update_window.title("성적 업데이트")
        self._center_window(update_window, 500, 350)
        update_window.configure(bg="#D3D3D3")

        form_frame = tk.Frame(update_window, bg="#D3D3D3")
        form_frame.pack(pady=20)

        label_font = ("D2Coding", 14)
        entry_font = ("D2Coding", 14)
        entry_width = 30

        tk.Label(form_frame, text="학생 ID:", bg="#D3D3D3", font=label_font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        student_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        student_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="강좌 ID:", bg="#D3D3D3", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        course_id_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        course_id_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        tk.Label(form_frame, text="성적:", bg="#D3D3D3", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        grade_entry = tk.Entry(form_frame, font=entry_font, width=entry_width)
        grade_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        tk.Button(update_window, text="저장", command=save_update_grade, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), width=15, relief="raised", bd=3).pack(pady=20)

    def view_student_details(self):
        def search_student():
            student_id = student_id_entry.get().strip()
            if student_id:
                try:
                    records = self.search_student_by_id(student_id)
                    for row in tree.get_children():
                        tree.delete(row)
                    if not records:
                        messagebox.showinfo("정보", "해당 학생 ID에 대한 정보가 없습니다.", parent=view_window)
                    for i, record in enumerate(records):
                        tag = "evenrow" if i % 2 == 0 else "oddrow"
                        tree.insert(
                            "",
                            "end",
                            values=(
                                record[0], # student_id
                                record[1], # name
                                record[2], # email
                                record[3], # course_name
                                record[4]  # grade
                            ),
                            tags=(tag,)
                        )
                except Exception as e:
                    messagebox.showerror("오류", f"오류 발생: {e}", parent=view_window)
            else:
                messagebox.showwarning("입력 필요", "검색할 학생 ID를 입력해주세요.", parent=view_window)

        view_window = tk.Toplevel(self.root)
        view_window.title("학생 세부 정보")
        self._center_window(view_window, 850, 550) # 창 크기 조정
        view_window.configure(bg="#D3D3D3")

        # 검색 프레임
        search_frame = tk.Frame(view_window, bg="#D3D3D3")
        search_frame.pack(pady=10)

        tk.Label(search_frame, text="학생 ID 입력:", bg="#D3D3D3", font=("D2Coding", 14)).pack(side=tk.LEFT, padx=5)
        student_id_entry = tk.Entry(search_frame, font=("D2Coding", 14), width=20)
        student_id_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(search_frame, text="검색", command=search_student, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), relief="raised", bd=3).pack(side=tk.LEFT, padx=10)

        # Treeview 스타일 설정
        style = ttk.Style()
        style.theme_use("clam") # 'clam', 'alt', 'default', 'classic' 등 다양한 테마 시도 가능
        style.configure("Treeview.Heading", font=("D2Coding", 12, "bold"), background="#4682B4", foreground="white")
        style.configure("Treeview", font=("D2Coding", 11), rowheight=25)
        style.map("Treeview", background=[('selected', '#0074D9')])

        # 교차 행 색상 설정
        tree = ttk.Treeview(view_window, columns=("학생 ID", "이름", "이메일", "강좌", "성적"), show="headings")
        tree.tag_configure("oddrow", background="#E0E0E0")
        tree.tag_configure("evenrow", background="#FFFFFF")

        tree.heading("학생 ID", text="학생 ID")
        tree.heading("이름", text="이름")
        tree.heading("이메일", text="이메일")
        tree.heading("강좌", text="강좌")
        tree.heading("성적", text="성적")

        tree.column("학생 ID", width=100, anchor="center")
        tree.column("이름", width=120, anchor="center")
        tree.column("이메일", width=200, anchor="center")
        tree.column("강좌", width=180, anchor="center")
        tree.column("성적", width=80, anchor="center")

        # 스크롤바 추가
        vsb = ttk.Scrollbar(view_window, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(view_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.pack(fill="both", expand=True, padx=10, pady=10)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def view_course_details(self):
        def search_course():
            course_id = course_id_entry.get().strip()
            if course_id:
                try:
                    records = self.search_course_by_id(course_id)
                    for row in tree.get_children():
                        tree.delete(row)
                    if not records:
                        messagebox.showinfo("정보", "해당 강좌 ID에 대한 정보가 없습니다.", parent=view_window)
                    for i, record in enumerate(records):
                        tag = "evenrow" if i % 2 == 0 else "oddrow"
                        tree.insert("", "end", values=(record[0], record[1], record[2], record[3], record[4]), tags=(tag,))
                except Exception as e:
                    messagebox.showerror("오류", f"오류 발생: {e}", parent=view_window)
            else:
                messagebox.showwarning("입력 필요", "검색할 강좌 ID를 입력해주세요.", parent=view_window)

        view_window = tk.Toplevel(self.root)
        view_window.title("강좌 세부 정보")
        self._center_window(view_window, 850, 550) # 창 크기 조정
        view_window.configure(bg="#D3D3D3")

        # 검색 프레임
        search_frame = tk.Frame(view_window, bg="#D3D3D3")
        search_frame.pack(pady=10)

        tk.Label(search_frame, text="강좌 ID 입력:", bg="#D3D3D3", font=("D2Coding", 14)).pack(side=tk.LEFT, padx=5)
        course_id_entry = tk.Entry(search_frame, font=("D2Coding", 14), width=20)
        course_id_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(search_frame, text="검색", command=search_course, bg="#008080", fg="white",
                  font=("D2Coding", 14, "bold"), relief="raised", bd=3).pack(side=tk.LEFT, padx=10)

        # Treeview 스타일 설정
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("D2Coding", 12, "bold"), background="#4682B4", foreground="white")
        style.configure("Treeview", font=("D2Coding", 11), rowheight=25)
        style.map("Treeview", background=[('selected', '#0074D9')])

        # 교차 행 색상 설정
        tree = ttk.Treeview(view_window, columns=("강좌 ID", "강좌 이름", "학생 ID", "이름", "성적"), show="headings")
        tree.tag_configure("oddrow", background="#E0E0E0")
        tree.tag_configure("evenrow", background="#FFFFFF")

        tree.heading("강좌 ID", text="강좌 ID")
        tree.heading("강좌 이름", text="강좌 이름")
        tree.heading("학생 ID", text="학생 ID")
        tree.heading("이름", text="이름")
        tree.heading("성적", text="성적")

        tree.column("강좌 ID", width=100, anchor="center")
        tree.column("강좌 이름", width=180, anchor="center")
        tree.column("학생 ID", width=100, anchor="center")
        tree.column("이름", width=120, anchor="center")
        tree.column("성적", width=80, anchor="center")

        # 스크롤바 추가
        vsb = ttk.Scrollbar(view_window, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(view_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.pack(fill="both", expand=True, padx=10, pady=10)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    # 전체 학생 목록 보기
    def view_all_students_details(self):
        view_window = tk.Toplevel(self.root)
        view_window.title("전체 학생 목록")
        self._center_window(view_window, 850, 550)
        view_window.configure(bg="#D3D3D3")

        tk.Label(view_window, text="전체 학생 목록", bg="#D3D3D3", font=("D2Coding", 16, "bold")).pack(pady=10)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("D2Coding", 12, "bold"), background="#4682B4", foreground="white")
        style.configure("Treeview", font=("D2Coding", 11), rowheight=25)
        style.map("Treeview", background=[('selected', '#0074D9')])

        tree = ttk.Treeview(view_window, columns=("학생 ID", "이름", "이메일", "강좌", "성적"), show="headings")
        tree.tag_configure("oddrow", background="#E0E0E0")
        tree.tag_configure("evenrow", background="#FFFFFF")

        tree.heading("학생 ID", text="학생 ID")
        tree.heading("이름", text="이름")
        tree.heading("이메일", text="이메일")
        tree.heading("강좌", text="강좌")
        tree.heading("성적", text="성적")

        tree.column("학생 ID", width=100, anchor="center")
        tree.column("이름", width=120, anchor="center")
        tree.column("이메일", width=200, anchor="center")
        tree.column("강좌", width=180, anchor="center")
        tree.column("성적", width=80, anchor="center")

        vsb = ttk.Scrollbar(view_window, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(view_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.pack(fill="both", expand=True, padx=10, pady=10)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        try:
            records = self.get_all_students_details()
            if not records:
                messagebox.showinfo("정보", "등록된 학생 정보가 없습니다.", parent=view_window)
            for i, record in enumerate(records):
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                tree.insert("", "end", values=record, tags=(tag,))
        except Exception as e:
            messagebox.showerror("오류", f"전체 학생 목록 조회 중 오류 발생: {e}", parent=view_window)

    # 전체 강좌 목록 보기
    def view_all_courses_details(self):
        view_window = tk.Toplevel(self.root)
        view_window.title("전체 강좌 목록")
        self._center_window(view_window, 850, 550)
        view_window.configure(bg="#D3D3D3")

        tk.Label(view_window, text="전체 강좌 목록", bg="#D3D3D3", font=("D2Coding", 16, "bold")).pack(pady=10)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("D2Coding", 12, "bold"), background="#4682B4", foreground="white")
        style.configure("Treeview", font=("D2Coding", 11), rowheight=25)
        style.map("Treeview", background=[('selected', '#0074D9')])

        tree = ttk.Treeview(view_window, columns=("강좌 ID", "강좌 이름", "학생 ID", "이름", "성적"), show="headings")
        tree.tag_configure("oddrow", background="#E0E0E0")
        tree.tag_configure("evenrow", background="#FFFFFF")

        tree.heading("강좌 ID", text="강좌 ID")
        tree.heading("강좌 이름", text="강좌 이름")
        tree.heading("학생 ID", text="학생 ID")
        tree.heading("이름", text="이름")
        tree.heading("성적", text="성적")

        tree.column("강좌 ID", width=100, anchor="center")
        tree.column("강좌 이름", width=180, anchor="center")
        tree.column("학생 ID", width=100, anchor="center")
        tree.column("이름", width=120, anchor="center")
        tree.column("성적", width=80, anchor="center")

        vsb = ttk.Scrollbar(view_window, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(view_window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.pack(fill="both", expand=True, padx=10, pady=10)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        try:
            records = self.get_all_courses_details()
            if not records:
                messagebox.showinfo("정보", "등록된 강좌 정보가 없습니다.", parent=view_window)
            for i, record in enumerate(records):
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                tree.insert("", "end", values=record, tags=(tag,))
        except Exception as e:
            messagebox.showerror("오류", f"전체 강좌 목록 조회 중 오류 발생: {e}", parent=view_window)

    def exit_program(self):
        if messagebox.askyesno("학생 관리 시스템", "정말로 종료하시겠습니까?", parent=self.root):
            self.root.quit()

    def open_update_interface(self):
        update_window = tk.Toplevel(self.root)
        update_window.title("업데이트 옵션")
        self._center_window(update_window, 400, 350)
        update_window.configure(bg="#FFD700")

        tk.Label(update_window, text="업데이트 옵션", bg="#FFD700", font=("D2Coding Bold", 18)).pack(pady=25)

        update_button_font = ("D2Coding", 14, "bold")
        update_button_width = 20
        update_button_pady = 12

        tk.Button(update_window, text="학생 업데이트", command=self.update_student, bg="#F0E68C",
                  font=update_button_font, width=update_button_width, relief="raised", bd=3).pack(pady=update_button_pady)
        tk.Button(update_window, text="강좌 업데이트", command=self.update_course, bg="#F0E68C",
                  font=update_button_font, width=update_button_width, relief="raised", bd=3).pack(pady=update_button_pady)
        tk.Button(update_window, text="성적 업데이트", command=self.update_grade, bg="#F0E68C",
                  font=update_button_font, width=update_button_width, relief="raised", bd=3).pack(pady=update_button_pady)

    def open_view_interface(self):
        view_window = tk.Toplevel(self.root)
        view_window.title("보기 옵션")
        self._center_window(view_window, 400, 350)
        view_window.configure(bg="#0074D9")

        tk.Label(view_window, text="보기 옵션", bg="#0074D9", fg="white", font=("D2Coding Bold", 18)).pack(pady=25)

        view_button_font = ("D2Coding", 14, "bold")
        view_button_width = 25
        view_button_pady = 12

        tk.Button(view_window, text="학생 세부 정보 보기", command=self.view_student_details, bg="#4682B4", fg="white",
                  font=view_button_font, width=view_button_width, relief="raised", bd=3).pack(pady=view_button_pady)
        tk.Button(view_window, text="강좌 세부 정보 보기", command=self.view_course_details, bg="#4682B4", fg="white",
                  font=view_button_font, width=view_button_width, relief="raised", bd=3).pack(pady=view_button_pady)
        tk.Button(view_window, text="전체 데이터 보기", command=self.open_all_data_view, bg="#4682B4", fg="white",
                  font=view_button_font, width=view_button_width, relief="raised", bd=3).pack(pady=view_button_pady)

    def open_all_data_view(self):
        all_data_window = tk.Toplevel(self.root)
        all_data_window.title("전체 데이터 보기 옵션")
        self._center_window(all_data_window, 400, 300)
        all_data_window.configure(bg="#ADD8E6") # 연한 파란색 배경

        tk.Label(all_data_window, text="전체 데이터 보기", bg="#ADD8E6", font=("D2Coding Bold", 18)).pack(pady=25)

        all_data_button_font = ("D2Coding", 14, "bold")
        all_data_button_width = 25
        all_data_button_pady = 12

        tk.Button(all_data_window, text="전체 학생 목록", command=self.view_all_students_details, bg="#87CEEB", fg="white",
                  font=all_data_button_font, width=all_data_button_width, relief="raised", bd=3).pack(pady=all_data_button_pady)
        tk.Button(all_data_window, text="전체 강좌 목록", command=self.view_all_courses_details, bg="#87CEEB", fg="white",
                  font=all_data_button_font, width=all_data_button_width, relief="raised", bd=3).pack(pady=all_data_button_pady)


if __name__ == "__main__":
    root = tk.Tk()
    obj = StudentManagementSystem(root)
    root.mainloop()