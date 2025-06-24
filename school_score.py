"""
school_score.py 학생 성적 관리 시스탬 Ver 1.0_250624
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3  # SQLite3 모듈 임포트


class SimpleStudentSystem:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("학생 결과 시스템")  # 제목 한국어화
        self.root.geometry("600x650")
        self.db = None
        self.cursor = None

        # 데이터베이스 연결 시도 및 성공 여부 확인
        if not self.connect_database():
            # 연결 실패 시 root 창을 파괴하고 __init__을 종료하여 더 이상 UI 작업을 하지 않도록 함
            self.root.destroy()
            return

        # 데이터베이스 연결 성공 시에만 인터페이스 생성
        self.create_interface()

    def connect_database(self):
        """
        SQLite3 데이터베이스에 연결합니다.
        데이터베이스 파일이 없으면 자동으로 생성됩니다.
        성공 시 True, 실패 시 False를 반환합니다.
        """
        try:
            # student_system.db 라는 파일로 데이터베이스 연결
            # 이 라인에서 student_system.db 파일이 없으면 자동으로 생성됩니다.
            self.db = sqlite3.connect('./school/student_system.db')
            self.cursor = self.db.cursor()
            # 외래 키 제약 조건 활성화 (기본적으로 비활성화되어 있음)
            self.cursor.execute('PRAGMA foreign_keys = ON;')
            print("데이터베이스에 성공적으로 연결되었습니다.")
            self.create_tables()
            return True  # 연결 및 테이블 생성 성공
        except sqlite3.Error as err:
            messagebox.showerror("데이터베이스 오류", f"데이터베이스 연결 실패: {err}\n"
                                              "파일 권한 문제 또는 디스크 공간 부족이 있을 수 있습니다.")
            return False  # 연결 실패
        except Exception as e:
            messagebox.showerror("일반 오류", f"데이터베이스 연결 중 예상치 못한 오류 발생: {e}")
            return False  # 기타 오류 발생

    def create_tables(self):
        """
        'students' 및 'marks' 테이블을 생성합니다(존재하지 않는 경우).
        """
        if not self.db or not self.cursor:
            # db 또는 cursor가 None인 경우, 연결 실패로 간주하고 테이블 생성 시도 안 함
            return

        try:
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, -- SQLite의 자동 증가 기본 키
                    name TEXT,
                    class TEXT,
                    roll_no TEXT UNIQUE
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS marks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER,
                    subject TEXT,
                    marks INTEGER,
                    skill TEXT,
                    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
                )
            """)
            self.db.commit()
            print("테이블 확인/생성 완료.")
        except sqlite3.Error as err:
            messagebox.showerror("데이터베이스 오류", f"테이블 생성 실패: {err}")
            self.db = None  # 테이블 생성 실패 시 DB 연결을 무효화
        except Exception as e:
            messagebox.showerror("일반 오류", f"테이블 생성 중 예상치 못한 오류 발생: {e}")
            self.db = None  # 테이블 생성 실패 시 DB 연결을 무효화

    def create_interface(self):
        """
        탭이 있는 메인 Tkinter GUI를 설정합니다.
        """
        title = tk.Label(self.root, text="학생 성적 관리 시스템",
                         font=("D2Coding", 20, "bold"), bg="lightblue")
        title.pack(fill=tk.X, pady=10)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.create_add_student_tab()
        self.create_add_marks_tab()
        self.create_view_results_tab()
        self.create_search_tab()

        # 모든 탭과 위젯이 생성된 후에 콤보박스 데이터 로드
        self.load_students_combo_and_results_data()

    def create_add_student_tab(self):
        """
        학생 정보 관리를 위한 '학생 추가' 탭을 생성합니다.
        """
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="학생 추가")  # 탭 이름 한국어화

        tk.Label(frame, text="새 학생 추가", font=("D2Coding", 16, "bold")).pack(pady=20)  # 텍스트 한국어화

        tk.Label(frame, text="학생 이름:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.name_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.name_entry.pack(pady=5)

        tk.Label(frame, text="반:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.class_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.class_entry.pack(pady=5)

        tk.Label(frame, text="학번:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.roll_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.roll_entry.pack(pady=5)

        tk.Button(frame, text="학생 추가", command=self.add_student,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 12, "bold"), bg="green", fg="white", width=28).pack(pady=10)

        tk.Button(frame, text="선택한 학생 정보 업데이트", command=self.update_student,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 10), bg="orange", fg="white", width=32).pack(pady=(5, 2))

        tk.Button(frame, text="선택한 학생 삭제", command=self.delete_student,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 10), bg="red", fg="white", width=32).pack(pady=(2, 10))

        tk.Label(frame, text="모든 학생:", font=("D2Coding", 14, "bold")).pack(pady=(10, 5))  # 텍스트 한국어화
        self.students_listbox = tk.Listbox(frame, height=10, font=("D2Coding", 10))
        self.students_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.refresh_students_list()

    def add_student(self):
        """
        데이터베이스에 새 학생 기록을 추가합니다.
        """
        name = self.name_entry.get().strip()
        class_name = self.class_entry.get().strip()
        roll_no = self.roll_entry.get().strip()

        if not name or not class_name or not roll_no:
            messagebox.showerror("오류", "모든 필드를 채워주세요!")  # 메시지 한국어화
            return

        try:
            # SQLite는 매개변수 바인딩에 '?'를 사용합니다.
            query = "INSERT INTO students (name, class, roll_no) VALUES (?, ?, ?)"
            self.cursor.execute(query, (name, class_name, roll_no))
            self.db.commit()
            messagebox.showinfo("성공", "학생이 성공적으로 추가되었습니다!")  # 메시지 한국어화
            self.name_entry.delete(0, tk.END)
            self.class_entry.delete(0, tk.END)
            self.roll_entry.delete(0, tk.END)
            self.refresh_students_list()
        except sqlite3.IntegrityError:
            messagebox.showerror("오류", "학번이 이미 존재합니다! 고유한 학번을 사용해주세요.")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"학생 추가 실패: {e}")  # 메시지 한국어화

    def update_student(self):
        """
        데이터베이스에서 기존 학생 기록을 업데이트합니다.
        """
        selected = self.students_listbox.curselection()
        if not selected:
            messagebox.showerror("오류", "업데이트할 학생을 목록에서 선택해주세요.")  # 메시지 한국어화
            return

        selected_text = self.students_listbox.get(selected[0])
        try:
            student_id = int(selected_text.split("ID: ")[1].split(" |")[0])
        except (ValueError, IndexError):
            messagebox.showerror("오류", "선택한 항목에서 학생 ID를 구문 분석할 수 없습니다.")  # 메시지 한국어화
            return

        name = self.name_entry.get().strip()
        class_name = self.class_entry.get().strip()
        roll_no = self.roll_entry.get().strip()

        if not name or not class_name or not roll_no:
            messagebox.showerror("오류", "업데이트를 위해 모든 필드를 채워주세요!")  # 메시지 한국어화
            return

        try:
            query = "UPDATE students SET name = ?, class = ?, roll_no = ? WHERE id = ?"
            self.cursor.execute(query, (name, class_name, roll_no, student_id))
            self.db.commit()
            if self.cursor.rowcount > 0:
                messagebox.showinfo("성공", "학생이 성공적으로 업데이트되었습니다!")  # 메시지 한국어화
                self.refresh_students_list()
                self.name_entry.delete(0, tk.END)
                self.class_entry.delete(0, tk.END)
                self.roll_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("경고", "업데이트할 선택된 ID를 가진 학생을 찾을 수 없습니다.")  # 메시지 한국어화
        except sqlite3.IntegrityError:
            messagebox.showerror("오류", "새 학번이 다른 학생에게 이미 존재합니다!")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"학생 업데이트 실패: {e}")  # 메시지 한국어화

    def delete_student(self):
        """
        데이터베이스에서 선택한 학생 기록과 모든 관련 성적을 삭제합니다.
        """
        selected = self.students_listbox.curselection()
        if not selected:
            messagebox.showerror("오류", "삭제할 학생을 목록에서 선택해주세요.")  # 메시지 한국어화
            return

        confirm = messagebox.askyesno("삭제 확인", "이 학생과 모든 성적을 정말 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")  # 메시지 한국어화
        if not confirm:
            return

        selected_text = self.students_listbox.get(selected[0])
        try:
            student_id = int(selected_text.split("ID: ")[1].split(" |")[0])
        except (ValueError, IndexError):
            messagebox.showerror("오류", "삭제를 위해 선택한 항목에서 학생 ID를 구문 분석할 수 없습니다.")  # 메시지 한국어화
            return

        try:
            query = "DELETE FROM students WHERE id = ?"
            self.cursor.execute(query, (student_id,))
            self.db.commit()
            if self.cursor.rowcount > 0:
                messagebox.showinfo("성공", "학생과 관련된 성적이 성공적으로 삭제되었습니다!")  # 메시지 한국어화
                self.refresh_students_list()
                self.name_entry.delete(0, tk.END)
                self.class_entry.delete(0, tk.END)
                self.roll_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("경고", "삭제할 선택된 ID를 가진 학생을 찾을 수 없습니다.")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"학생 삭제 실패: {e}")  # 메시지 한국어화

    def refresh_students_list(self):
        """
        '학생 추가' 탭에 있는 모든 학생을 표시하는 목록 상자를 새로 고칩니다.
        """
        self.students_listbox.delete(0, tk.END)
        try:
            self.cursor.execute("SELECT id, name, class, roll_no FROM students ORDER BY name ASC")
            students = self.cursor.fetchall()
            if not students:
                self.students_listbox.insert(tk.END, "아직 추가된 학생이 없습니다.")  # 메시지 한국어화
            for student in students:
                student_info = f"ID: {student[0]} | {student[1]} | 반: {student[2]} | 학번: {student[3]}"  # 텍스트 한국어화
                self.students_listbox.insert(tk.END, student_info)
        except Exception as e:
            messagebox.showerror("오류", f"학생 목록 불러오기 실패: {e}")  # 메시지 한국어화

    def create_add_marks_tab(self):
        """
        학생 성적 관리를 위한 '성적 추가' 탭을 생성합니다.
        """
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="성적 추가")  # 탭 이름 한국어화

        tk.Label(frame, text="학생 성적 추가 / 업데이트", font=("D2Coding", 16, "bold")).pack(pady=10)  # 텍스트 한국어화

        tk.Label(frame, text="학생 선택:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.student_combo = ttk.Combobox(frame, font=("D2Coding", 12), width=40, state="readonly")
        self.student_combo.pack(pady=5)
        self.student_combo.bind("<<ComboboxSelected>>", self.on_student_combo_select)
        # self.load_students_combo() # 이 호출은 create_interface의 load_students_combo_and_results_data로 이동됨

        tk.Label(frame, text="선택한 학생의 기존 과목:", font=("D2Coding", 12)).pack(pady=(10, 5))  # 텍스트 한국어화
        self.subjects_listbox = tk.Listbox(frame, height=5, font=("D2Coding", 10))
        self.subjects_listbox.pack(pady=5, fill=tk.X, padx=150)
        self.subjects_listbox.bind("<<ListboxSelect>>", self.load_selected_subject_data)

        tk.Label(frame, text="과목 이름:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.subject_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.subject_entry.pack(pady=5)

        tk.Label(frame, text="점수 (0-100):", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.marks_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.marks_entry.pack(pady=5)

        tk.Label(frame, text="기술 (예: Python, Java):", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.skill_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.skill_entry.pack(pady=5)

        tk.Button(frame, text="새 성적 추가", command=self.add_marks,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 12, "bold"), bg="blue", fg="white", width=27).pack(pady=5)

        tk.Button(frame, text="선택한 성적 업데이트", command=self.update_marks,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 10), bg="orange", fg="white", width=31).pack(pady=2)

        tk.Button(frame, text="선택한 성적 삭제", command=self.delete_marks,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 10), bg="red", fg="white", width=31).pack(pady=2)

    # load_students_combo 함수는 이제 사용되지 않음. load_students_combo_and_results_data로 통합됨
    # def load_students_combo(self):
    #     pass

    def on_student_combo_select(self, event):
        """
        콤보박스에서 학생이 선택될 때의 이벤트 핸들러입니다.
        선택한 학생의 과목을 로드합니다.
        """
        self.load_student_subjects()
        self.subject_entry.delete(0, tk.END)
        self.marks_entry.delete(0, tk.END)
        self.skill_entry.delete(0, tk.END)

    def load_student_subjects(self):
        """
        현재 선택된 학생의 과목 및 세부 정보를 로드합니다.
        """
        self.subjects_listbox.delete(0, tk.END)
        selected_student = self.student_combo.get()
        if not selected_student or "ID: " not in selected_student:
            return

        try:
            student_id = int(selected_student.split("ID: ")[1])
            self.cursor.execute("SELECT subject FROM marks WHERE student_id = ? ORDER BY subject ASC", (student_id,))
            subjects = self.cursor.fetchall()
            if not subjects:
                self.subjects_listbox.insert(tk.END, "이 학생에게 아직 성적이 추가되지 않았습니다.")  # 메시지 한국어화
            for subject in subjects:
                self.subjects_listbox.insert(tk.END, subject[0])
        except Exception as e:
            messagebox.showerror("오류", f"학생 과목 불러오기 실패: {e}")  # 메시지 한국어화

    def load_selected_subject_data(self, event):
        """
        선택한 과목의 세부 정보(과목, 점수, 기술)를 편집 필드에 로드합니다.
        """
        try:
            if not self.subjects_listbox.curselection():
                return

            subject_name = self.subjects_listbox.get(self.subjects_listbox.curselection())
            selected_student = self.student_combo.get()
            if not selected_student or "ID: " not in selected_student:
                return

            student_id = int(selected_student.split("ID: ")[1])
            self.cursor.execute("SELECT subject, marks, skill FROM marks WHERE student_id = ? AND subject = ?",
                                (student_id, subject_name))
            data = self.cursor.fetchone()
            if data:
                self.subject_entry.delete(0, tk.END)
                self.marks_entry.delete(0, tk.END)
                self.skill_entry.delete(0, tk.END)

                self.subject_entry.insert(0, data[0])
                self.marks_entry.insert(0, data[1])
                self.skill_entry.insert(0, data[2])
        except IndexError:
            pass
        except Exception as e:
            messagebox.showerror("오류", f"과목 데이터 불러오기 실패: {e}")  # 메시지 한국어화

    def add_marks(self):
        """
        선택한 학생에게 새 성적을 추가합니다. 중복 과목 입력을 방지합니다.
        """
        selected_student = self.student_combo.get()
        if not selected_student or "ID: " not in selected_student:
            messagebox.showerror("오류", "먼저 드롭다운에서 학생을 선택해주세요!")  # 메시지 한국어화
            return

        subject = self.subject_entry.get().strip()
        marks_text = self.marks_entry.get().strip()
        skill = self.skill_entry.get().strip()

        if not subject or not marks_text or not skill:
            messagebox.showerror("오류", "성적 입력을 위해 모든 필드를 채워주세요!")  # 메시지 한국어화
            return

        try:
            marks = int(marks_text)
            if not (0 <= marks <= 100):
                messagebox.showerror("오류", "점수는 0에서 100 사이의 정수여야 합니다.")  # 메시지 한국어화
                return

            student_id = int(selected_student.split("ID: ")[1])

            self.cursor.execute("SELECT COUNT(*) FROM marks WHERE student_id = ? AND subject = ?",
                                (student_id, subject))
            if self.cursor.fetchone()[0] > 0:
                messagebox.showerror("오류", f"이 학생에게 '{subject}' 과목의 성적이 이미 존재합니다. '업데이트'를 사용하여 변경하세요.")  # 메시지 한국어화
                return

            query = "INSERT INTO marks (student_id, subject, marks, skill) VALUES (?, ?, ?, ?)"
            self.cursor.execute(query, (student_id, subject, marks, skill))
            self.db.commit()
            messagebox.showinfo("성공", f"{subject} 과목 성적이 성공적으로 추가되었습니다!")  # 메시지 한국어화
            self.subject_entry.delete(0, tk.END)
            self.marks_entry.delete(0, tk.END)
            self.skill_entry.delete(0, tk.END)
            self.load_student_subjects()
        except ValueError:
            messagebox.showerror("입력 오류", "점수는 유효한 숫자여야 합니다.")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"성적 추가 실패: {e}")  # 메시지 한국어화

    def update_marks(self):
        """
        선택한 학생의 선택한 과목에 대한 성적을 업데이트합니다.
        """
        selected_student = self.student_combo.get()
        if not selected_student or "ID: " not in selected_student:
            messagebox.showerror("오류", "먼저 학생을 선택해주세요!")  # 메시지 한국어화
            return

        selected_subject_index = self.subjects_listbox.curselection()
        if not selected_subject_index:
            messagebox.showerror("오류", "업데이트할 과목을 목록에서 선택해주세요.")  # 메시지 한국어화
            return

        old_subject_name = self.subjects_listbox.get(selected_subject_index[0])

        new_subject = self.subject_entry.get().strip()
        new_marks_text = self.marks_entry.get().strip()
        new_skill = self.skill_entry.get().strip()

        if not new_subject or not new_marks_text or not new_skill:
            messagebox.showerror("오류", "업데이트를 위해 모든 필드(과목, 점수, 기술)를 채워야 합니다.")  # 메시지 한국어화
            return

        try:
            new_marks = int(new_marks_text)
            if not (0 <= new_marks <= 100):
                messagebox.showerror("오류", "점수는 0에서 100 사이의 정수여야 합니다.")  # 메시지 한국어화
                return

            student_id = int(selected_student.split("ID: ")[1])

            if old_subject_name != new_subject:
                self.cursor.execute("SELECT COUNT(*) FROM marks WHERE student_id = ? AND subject = ?",
                                    (student_id, new_subject))
                if self.cursor.fetchone()[0] > 0:
                    messagebox.showerror("오류", f"과목 '{new_subject}'이(가) 이 학생에게 이미 존재합니다. 고유한 과목 이름을 선택하거나 기존 과목을 업데이트하세요.")  # 메시지 한국어화
                    return

            query = "UPDATE marks SET subject = ?, marks = ?, skill = ? WHERE student_id = ? AND subject = ?"
            self.cursor.execute(query, (new_subject, new_marks, new_skill, student_id, old_subject_name))
            self.db.commit()
            if self.cursor.rowcount > 0:
                messagebox.showinfo("성공", f"'{old_subject_name}' 과목 성적이 성공적으로 업데이트되었습니다!")  # 메시지 한국어화
                self.load_student_subjects()
                self.subject_entry.delete(0, tk.END)
                self.marks_entry.delete(0, tk.END)
                self.skill_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("경고", "업데이트할 선택된 과목에 대한 성적을 찾을 수 없습니다.")  # 메시지 한국어화
        except ValueError:
            messagebox.showerror("입력 오류", "점수는 유효한 숫자여야 합니다.")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"성적 업데이트 실패: {e}")  # 메시지 한국어화

    def delete_marks(self):
        """
        특정 학생에 대해 선택된 성적을 삭제합니다.
        """
        selected_student = self.student_combo.get()
        if not selected_student or "ID: " not in selected_student:
            messagebox.showerror("오류", "먼저 학생을 선택해주세요!")  # 메시지 한국어화
            return

        selected_subject_index = self.subjects_listbox.curselection()
        if not selected_subject_index:
            messagebox.showerror("오류", "삭제할 과목을 목록에서 선택해주세요.")  # 메시지 한국어화
            return

        subject_to_delete = self.subjects_listbox.get(selected_subject_index[0])
        confirm = messagebox.askyesno("삭제 확인", f"'{subject_to_delete}' 과목 성적을 정말 삭제하시겠습니까?")  # 메시지 한국어화
        if not confirm:
            return

        try:
            student_id = int(selected_student.split("ID: ")[1])
            query = "DELETE FROM marks WHERE student_id = ? AND subject = ?"
            self.cursor.execute(query, (student_id, subject_to_delete))
            self.db.commit()
            if self.cursor.rowcount > 0:
                messagebox.showinfo("삭제됨", f"'{subject_to_delete}' 과목 성적이 성공적으로 삭제되었습니다.")  # 메시지 한국어화
                self.load_student_subjects()
                self.subject_entry.delete(0, tk.END)
                self.marks_entry.delete(0, tk.END)
                self.skill_entry.delete(0, tk.END)
            else:
                messagebox.showwarning("경고", "삭제할 선택된 과목에 대한 성적을 찾을 수 없습니다.")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"성적 삭제 실패: {e}")  # 메시지 한국어화

    def create_view_results_tab(self):
        """
        학생 성적표를 표시하기 위한 '결과 보기' 탭을 생성합니다.
        """
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="결과 보기")  # 탭 이름 한국어화

        tk.Label(frame, text="학생 결과 보기", font=("D2Coding", 16, "bold")).pack(pady=20)  # 텍스트 한국어화

        tk.Label(frame, text="학생 선택:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.result_combo = ttk.Combobox(frame, font=("D2Coding", 12), width=40, state="readonly")
        self.result_combo.pack(pady=5)
        # self.load_result_combo() # 이 호출은 create_interface의 load_students_combo_and_results_data로 이동됨

        tk.Button(frame, text="결과 보기", command=self.view_result,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 12, "bold"), bg="orange", fg="white", width=42).pack(pady=10)

        self.result_text = tk.Text(frame, height=20, width=75, font=("D2Coding", 11))
        self.result_text.pack(pady=20, padx=60, fill=tk.BOTH, expand=True)

    # load_result_combo 함수는 이제 사용되지 않음. load_students_combo_and_results_data로 통합됨
    # def load_result_combo(self):
    #     pass

    def load_students_combo_and_results_data(self):
        """
        학생 콤보박스와 결과 콤보박스를 모두 학생 데이터로 채웁니다.
        모든 UI 위젯이 생성된 후에 호출됩니다.
        """
        try:
            self.cursor.execute("SELECT id, name, roll_no FROM students ORDER BY name ASC")
            students = self.cursor.fetchall()
            student_list = [f"{s[1]} (학번: {s[2]}) - ID: {s[0]}" for s in students]

            self.student_combo['values'] = student_list
            self.result_combo['values'] = student_list

            if student_list:
                self.student_combo.set(student_list[0])
                self.result_combo.set(student_list[0])
                # '성적 추가' 탭에서 첫 번째 학생에 대한 과목 로드를 수동으로 트리거
                self.on_student_combo_select(None)
            else:
                self.student_combo.set("사용 가능한 학생이 없습니다")  # 메시지 한국어화
                self.result_combo.set("사용 가능한 학생이 없습니다")  # 메시지 한국어화
        except Exception as e:
            messagebox.showerror("오류", f"학생 콤보박스 데이터 불러오기 실패: {e}")  # 메시지 한국어화

    def view_result(self):
        """
        선택한 학생의 성적표를 생성하고 표시합니다.
        """
        selected_student = self.result_combo.get()
        if not selected_student or "ID: " not in selected_student:
            messagebox.showerror("오류", "결과를 보려면 드롭다운에서 학생을 선택해주세요!")  # 메시지 한국어화
            return

        try:
            student_id = int(selected_student.split("ID: ")[1])

            self.cursor.execute("SELECT name, class, roll_no FROM students WHERE id = ?", (student_id,))
            student_data = self.cursor.fetchone()

            if not student_data:
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, "학생 데이터를 찾을 수 없습니다. (목록에서 선택한 경우 발생해서는 안 됩니다.)")  # 메시지 한국어화
                return

            self.cursor.execute("SELECT subject, marks, skill FROM marks WHERE student_id = ? ORDER BY subject ASC", (student_id,))
            marks_data = self.cursor.fetchall()

            self.result_text.delete(1.0, tk.END)

            if not marks_data:
                self.result_text.insert(tk.END, f"{student_data[0]} (학번: {student_data[2]}) 학생에게 기록된 성적이 없습니다.")  # 메시지 한국어화
                return

            result_text = f"학생 성적표\n{'=' * 50}\n\n"  # 텍스트 한국어화
            result_text += f"이름: {student_data[0]}\n반: {student_data[1]}\n학번: {student_data[2]}\n\n"  # 텍스트 한국어화
            result_text += f"성적 및 기술:\n{'-' * 40}\n"  # 텍스트 한국어화

            total_marks_obtained = 0
            total_possible_marks = 0

            for subject, mark, skill in marks_data:
                result_text += f"{subject}: {mark}/100 | 기술: {skill}\n"  # 텍스트 한국어화
                total_marks_obtained += mark
                total_possible_marks += 100

            percentage = (total_marks_obtained / total_possible_marks) * 100 if total_possible_marks > 0 else 0

            grade = ""
            if percentage >= 90:
                grade = "A+"
            elif percentage >= 80:
                grade = "A"
            elif percentage >= 70:
                grade = "B"
            elif percentage >= 60:
                grade = "C"
            elif percentage >= 50:
                grade = "D"
            else:
                grade = "F (불합격)"  # 텍스트 한국어화

            academic_performance = ""
            if percentage >= 75:
                academic_performance = "우수 학업 성과"  # 텍스트 한국어화
            elif percentage >= 60:
                academic_performance = "양호 학업 성과"  # 텍스트 한국어화
            elif percentage >= 50:
                academic_performance = "평균 학업 성과"  # 텍스트 한국어화
            else:
                academic_performance = "상당한 개선 필요"  # 텍스트 한국어화

            result_text += f"\n{'-' * 40}\n"
            result_text += f"총 획득 점수: {total_marks_obtained}/{total_possible_marks}\n"  # 텍스트 한국어화
            result_text += f"전체 백분율: {percentage:.2f}%\n"  # 텍스트 한국어화
            result_text += f"최종 등급: {grade}\n"  # 텍스트 한국어화
            result_text += f"학업 성과: {academic_performance}\n"  # 텍스트 한국어화
            result_text += f"{'=' * 50}\n"

            self.result_text.insert(tk.END, result_text)
        except Exception as e:
            messagebox.showerror("오류", f"결과 불러오기 실패: {e}")  # 메시지 한국어화

    def create_search_tab(self):
        """
        학생 기록 검색을 위한 '검색' 탭을 생성합니다.
        """
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="검색")  # 탭 이름 한국어화

        tk.Label(frame, text="학생 검색", font=("D2Coding", 16, "bold")).pack(pady=20)  # 텍스트 한국어화

        tk.Label(frame, text="학생 이름 또는 학번 입력:", font=("D2Coding", 12)).pack()  # 텍스트 한국어화
        self.search_entry = tk.Entry(frame, font=("D2Coding", 12), width=30)
        self.search_entry.pack(pady=10)

        tk.Button(frame, text="검색", command=self.search_student,  # 버튼 텍스트 한국어화
                  font=("D2Coding", 12, "bold"), bg="purple", fg="white", width=28).pack(pady=10)

        tk.Label(frame, text="검색 결과:", font=("D2Coding", 12, "bold")).pack(pady=(20, 10))  # 텍스트 한국어화
        self.search_listbox = tk.Listbox(frame, height=15, font=("D2Coding", 10))
        self.search_listbox.pack(fill=tk.BOTH, expand=True, padx=60, pady=10)
        self.search_listbox.bind("<Double-Button-1>", self.on_search_result_double_click)

    def search_student(self):
        """
        이름 또는 학번으로 학생을 검색하고 결과를 표시합니다.
        """
        search_term = self.search_entry.get().strip()

        if not search_term:
            messagebox.showerror("오류", "검색어를 입력해주세요 (이름 또는 학번)!")  # 메시지 한국어화
            return

        try:
            query = "SELECT id, name, class, roll_no FROM students WHERE name LIKE ? OR roll_no LIKE ? ORDER BY name ASC"
            search_pattern = f"%{search_term}%"
            self.cursor.execute(query, (search_pattern, search_pattern))
            results = self.cursor.fetchall()

            self.search_listbox.delete(0, tk.END)

            if not results:
                self.search_listbox.insert(tk.END, "검색어와 일치하는 학생을 찾을 수 없습니다.")  # 메시지 한국어화
            else:
                for student in results:
                    student_info = f"ID: {student[0]} | {student[1]} | 반: {student[2]} | 학번: {student[3]}"  # 텍스트 한국어화
                    self.search_listbox.insert(tk.END, student_info)
        except Exception as e:
            messagebox.showerror("오류", f"검색 실패: {e}")  # 메시지 한국어화

    def on_search_result_double_click(self, event):
        """
        검색 결과 목록 상자에서 더블 클릭 이벤트를 처리합니다.
        '결과 보기' 탭으로 전환하고 선택한 학생의 결과를 표시합니다.
        """
        try:
            selected_index = self.search_listbox.curselection()
            if not selected_index:
                return

            selected_text = self.search_listbox.get(selected_index[0])
            student_id = int(selected_text.split("ID: ")[1].split(" |")[0])

            student_combo_values = self.result_combo['values']
            found_item = None
            for item in student_combo_values:
                if f"ID: {student_id}" in item:
                    found_item = item
                    break

            if found_item:
                self.notebook.select(self.create_view_results_tab_index())
                self.result_combo.set(found_item)
                self.view_result()
            else:
                messagebox.showwarning("찾을 수 없음", "결과 보기 드롭다운에서 학생을 찾을 수 없습니다.")  # 메시지 한국어화

        except Exception as e:
            messagebox.showerror("오류", f"검색 결과에서 결과 보기 실패: {e}")  # 메시지 한국어화

    def create_view_results_tab_index(self):
        """
        '결과 보기' 탭의 인덱스를 가져오는 헬퍼입니다.
        """
        for i, tab_id in enumerate(self.notebook.tabs()):
            if self.notebook.tab(tab_id, "text") == "결과 보기":  # 탭 이름 한국어화
                return tab_id
        return None

    def run(self):
        """
        Tkinter 이벤트 루프를 시작합니다.
        """
        self.root.mainloop()

    def __del__(self):
        """
        애플리케이션 종료 시 데이터베이스 연결이 닫히도록 합니다.
        """
        if self.cursor:
            self.cursor.close()
        if self.db:
            self.db.close()
        print("데이터베이스 연결이 닫혔습니다.")


if __name__ == "__main__":
    try:
        app = SimpleStudentSystem()
        if app.db:
            app.run()
    except Exception as e:
        print(f"애플리케이션 시작 중 처리되지 않은 오류 발생: {e}")
        input("Enter 키를 눌러 종료...")  # 메시지 한국어화
