"""
studentmanagesystem.py 학생 관리 시스템  Ver 1.0_250313
"""
import sys
import sqlite3
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem, QMessageBox, QStatusBar
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from datetime import datetime

DB_FILE = "./school/studentsdb.db"
 
VALID_PHONE_LENGTH = [11, 12]

def create_connection():
    conn = sqlite3.connect(DB_FILE)
    return conn

def initialize_db():
    conn = create_connection()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            birthYear TEXT NOT NULL,
            grade TEXT NOT NULL,
            phone TEXT NOT NULL UNIQUE
        )
    ''')
    conn.commit()
    conn.close()

def load_students():
    conn = create_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, birthYear, grade, phone FROM students")
    students = cursor.fetchall()
    conn.close()
    return students

def save_student(student):
    conn = create_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO students (name, birthYear, grade, phone) VALUES (?, ?, ?, ?)",
                   (student['name'], student['birthYear'], student['grade'], student['phone']))
    conn.commit()
    conn.close()

def update_student_in_db(student, phone):
    conn = create_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE students SET name = ?, birthYear = ?, grade = ?, phone = ? WHERE phone = ?",
                   (student['name'], student['birthYear'], student['grade'], student['phone'], phone))
    conn.commit()
    conn.close()

def delete_student_from_db(phone):
    conn = create_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM students WHERE phone = ?", (phone,))
    conn.commit()
    conn.close()

class StudentManagementApp(QMainWindow):
    def __init__(self):
        # Main window initialization
        super().__init__()
        self.setWindowTitle("학생 관리 시스템")
        self.setGeometry(125, 125, 700, 650)
        # Main widget initialization
        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)
        self.layout = QVBoxLayout(self.centralWidget)
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("학생 관리 시스템에 오신걸 환영합니다.", 5000)
        # Input Data area
        self.input_layout = QHBoxLayout()
        self.name_input = QLineEdit()
        self.birthYear_input = QLineEdit()
        self.grade_input = QLineEdit()
        self.phone_input = QLineEdit()
        self.input_layout.addWidget(QLabel("이름:"))
        self.input_layout.addWidget(self.name_input)
        self.input_layout.addWidget(QLabel("생년:"))
        self.input_layout.addWidget(self.birthYear_input)
        self.input_layout.addWidget(QLabel("학년:"))
        self.input_layout.addWidget(self.grade_input)
        self.input_layout.addWidget(QLabel("전화:"))
        self.input_layout.addWidget(self.phone_input)
        # Basic Button area
        self.button_layout = QHBoxLayout()
        self.add_button = QPushButton("학생 추가")
        self.delete_button = QPushButton("학생 삭제")
        self.update_button = QPushButton("학생 수정")
        self.button_layout.addWidget(self.add_button)
        self.button_layout.addWidget(self.delete_button)
        self.button_layout.addWidget(self.update_button)
        # Search and Sort layout area
        self.search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_button = QPushButton("검색")
        self.sort_button = QPushButton("정렬")
        self.reverse_sort_button = QPushButton("역정렬")
        self.search_button.setFixedSize(100, 30)
        self.sort_button.setFixedSize(100, 30)
        self.reverse_sort_button.setFixedSize(100, 30)
        self.search_layout.addWidget(QLabel("전화로 검색:"))
        self.search_layout.addWidget(self.search_input)
        self.search_layout.addWidget(self.search_button)
        self.search_layout.addWidget(self.sort_button)
        self.search_layout.addWidget(self.reverse_sort_button)
        # Student table area
        self.student_list = QTreeWidget()
        self.student_list.setColumnCount(4)
        self.student_list.setHeaderLabels(["Name", "birthYear", "Grade", "Phone"])
        # Add widgets to the main layout
        self.layout.addLayout(self.input_layout)
        self.layout.addLayout(self.button_layout)
        self.layout.addLayout(self.search_layout)
        self.layout.addWidget(self.student_list)
        # Connect the buttons to the functions
        self.add_button.clicked.connect(self.add_student)
        self.delete_button.clicked.connect(self.delete_student)
        self.update_button.clicked.connect(self.update_student)
        self.student_list.itemSelectionChanged.connect(self.select_student)
        self.search_button.clicked.connect(self.search_student)
        self.sort_button.clicked.connect(self.sort_students)
        self.reverse_sort_button.clicked.connect(self.reverse_sort_students)
        # Initialize the database
        initialize_db()
        # Update student list
        self.update_student_list()
        # Set font
        self.set_font()

    def set_font(self):
        font = QFont("맑은 고딕", 12)
        self.setFont(font)

    def reverse_sort_students(self):
        if self.student_list.topLevelItemCount() > 0:
            students = load_students()
            students.sort(key=lambda x: x[0], reverse=True)
            self.update_student_list(students)
            self.status_bar.showMessage("이름 역순으로 정렬합니다!", 5000)
        else:
            QMessageBox.warning(self, "오류", "정렬할 학생이 없습니다!")

    def sort_students(self):
        if self.student_list.topLevelItemCount() > 0:
            students = load_students()
            students.sort(key=lambda x: x[0])
            self.update_student_list(students)
            self.status_bar.showMessage("이름순으로 정렬합니다!", 5000)
        else:
            QMessageBox.warning(self, "오류", "정렬할 학생이 없습니다!")

    def select_student(self):
        selected = self.student_list.currentItem()
        if selected:
            self.name_input.setText(selected.text(0))
            self.birthYear_input.setText(selected.text(1))
            self.grade_input.setText(selected.text(2))
            self.phone_input.setText(selected.text(3))

    def update_student_list(self, students=None):
        self.student_list.clear()
        students = load_students() if students is None else students
        now = datetime.now()
        year = now.year
        for student in students:
            item = QTreeWidgetItem([student[0], student[1], student[2], student[3]])
            self.student_list.addTopLevelItem(item)

    def add_student(self):
        name = self.name_input.text().strip()
        birthYear = self.birthYear_input.text().strip()
        grade = self.grade_input.text().strip()
        phone = self.phone_input.text().strip()
        student = {"name": name, "birthYear": birthYear, "grade": grade, "phone": phone}
        if name and birthYear and grade and phone:
            if not self.check_OK(student):
                return
            if self.phone_exists(phone):
                QMessageBox.warning(self, "Error", "동일한 전화번호가 존재합니다!")
                return
            save_student(student)
            self.update_student_list()
            self.clear_inputs()
            self.status_bar.showMessage("자료가 성공적으로 추가되었습니다!", 5000)
        else:
            QMessageBox.warning(self, "Error", "Please fill all fields!")

    def delete_student(self):
        selected = self.student_list.currentItem()
        if selected:
            phone = selected.text(3)
            delete_student_from_db(phone)
            self.update_student_list()
            self.student_list.clearSelection()
            self.clear_inputs()
            self.status_bar.showMessage("성공적으로 삭제되었습니다!", 5000)
        else:
            QMessageBox.warning(self, "오류", "삭제할 학생을 선택하시오!")

    def update_student(self):
        selected = self.student_list.currentItem()
        if selected:
            phone = selected.text(3)
            name = self.name_input.text().strip()
            birthYear = self.birthYear_input.text().strip()
            grade = self.grade_input.text().strip()
            new_phone = self.phone_input.text().strip()
            student = {"name": name, "birthYear": birthYear, "grade": grade, "phone": new_phone}
            if name and birthYear and grade and new_phone:
                if not self.check_OK(student):
                    return
                update_student_in_db(student, phone)
                self.update_student_list()
                self.student_list.clearSelection()
                self.clear_inputs()
                self.status_bar.showMessage("자료가 성공적으로 수정되었습니다!", 5000)
            else:
                QMessageBox.warning(self, "오류", "모든 항목을 채우시오!")
        else:
            QMessageBox.warning(self, "오류", "수정할 학생을 선택하시오!")

    def search_student(self):
        student_phone = self.search_input.text().strip()
        if self.phone_exists(student_phone):
            students = load_students()
            for student in students:
                if student[3] == student_phone:
                    self.name_input.setText(student[0])
                    self.birthYear_input.setText(student[1])
                    self.grade_input.setText(student[2])
                    self.phone_input.setText(student[3])
                    item = self.student_list.findItems(student[3], Qt.MatchFlag.MatchExactly, 3)
                    self.student_list.scrollToItem(item[0])
                    self.student_list.setCurrentItem(item[0])
        elif student_phone == "":
            QMessageBox.warning(self, "오류", "검색할 전화번호흫 입력하세요!")
        else:
            QMessageBox.warning(self, "오류", "학생을 찾을 수 없습니다!")

    def phone_exists(self, phone):
        students = load_students()
        for student in students:
            if student[3] == phone:
                return True
        return False

    def check_OK(self, student):
        now = datetime.now()
        year = now.year
        if not student["birthYear"].isdigit():
            QMessageBox.warning(self, "오류", "생년은 숫자여야 합니다!")
            return False
        elif not int(student['birthYear']) <= year:
            QMessageBox.warning(self, "오류", "생년은 현재 연도보다 작아야 합니다!")
            return False
        if not student['grade'].isdigit():
            QMessageBox.warning(self, "오류", "학년은 숫자여야만 합니다.!")
            return False
        if not len(student['phone']) in VALID_PHONE_LENGTH:
            QMessageBox.warning(self, "오류", "전화번호는 11자리 수이여하 합니다!")
            return False
        return True

    def clear_inputs(self):
        self.name_input.clear()
        self.birthYear_input.clear()
        self.grade_input.clear()
        self.phone_input.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StudentManagementApp()
    window.show()
    sys.exit(app.exec_())