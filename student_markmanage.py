"""
student_markmanage.py 학생 성적 관리  Ver 1.0_251128
"""
import sys
import csv
import sqlite3
from pathlib import Path
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import (
    QDialog, QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QMessageBox, QComboBox, QSpinBox, QFileDialog, QDateEdit,
)
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon

# --- Database 구성 ---
DB_NAME = "student_results.db"


def get_connection():
    return sqlite3.connect(DB_NAME)


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    # 1. 학생 Table
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            class_name TEXT NOT NULL,
            section TEXT,
            dob TEXT
        )
        """
    )

    # 2. 과목 Table
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS subjects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            max_marks INTEGER NOT NULL
        )
        """
    )

    # 3. 성적 Table(학생 및 과목 연결)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS marks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            exam TEXT NOT NULL,
            marks_obtained REAL NOT NULL,
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
            FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE CASCADE
        )
        """
    )

    conn.commit()
    conn.close()


# --- 학생 CRUD 운영 ---
def add_student(roll_no, name, class_name, section, dob):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO students (roll_no, name, class_name, section, dob) VALUES (?, ?, ?, ?, ?)",
        (roll_no, name, class_name, section, dob),
    )
    conn.commit()
    conn.close()


def update_student(student_id, roll_no, name, class_name, section, dob):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE students
        SET roll_no = ?, name = ?, class_name = ?, section = ?, dob = ?
        WHERE id = ?
        """,
        (roll_no, name, class_name, section, dob, student_id),
    )
    conn.commit()
    conn.close()


def delete_student(student_id):
    conn = get_connection()
    cur = conn.cursor()
    # 관련 마크 삭제는 ON DELETE CASCADE로 처리됩니다.
    cur.execute("DELETE FROM students WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()


def get_students():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, roll_no, name, class_name, section, dob FROM students ORDER BY class_name, roll_no"
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_student_by_roll(roll_no):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, roll_no, name, class_name, section, dob FROM students WHERE roll_no = ?",
        (roll_no,),
    )
    row = cur.fetchone()
    conn.close()
    return row


def search_students(keyword):
    pattern = f"%{keyword}%"
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, roll_no, name, class_name, section, dob FROM students
        WHERE roll_no LIKE ? OR name LIKE ? OR class_name LIKE ?
        ORDER BY class_name, roll_no
        """,
        (pattern, pattern, pattern),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


# --- 주제 CRUD 작업 ---
def add_subject(code, name, max_marks):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO subjects (code, name, max_marks) VALUES (?, ?, ?)",
        (code, name, max_marks),
    )
    conn.commit()
    conn.close()


def update_subject(subject_id, code, name, max_marks):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE subjects
        SET code = ?, name = ?, max_marks = ?
        WHERE id = ?
        """,
        (code, name, max_marks, subject_id),
    )
    conn.commit()
    conn.close()


def delete_subject(subject_id):
    conn = get_connection()
    cur = conn.cursor()
    # 관련 점수 삭제는 ON DELETE CASCADE로 처리됩니다.
    cur.execute("DELETE FROM subjects WHERE id = ?", (subject_id,))
    conn.commit()
    conn.close()


def get_subjects():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, code, name, max_marks FROM subjects ORDER BY code"
    )
    rows = cur.fetchall()
    conn.close()
    return rows


# --- 성적 운영 ---
def set_marks(student_id, subject_id, exam, marks_obtained):
    conn = get_connection()
    cur = conn.cursor()
    # Upsert와 유사한 동작: 레코드가 있으면 업데이트하고, 그렇지 않으면 삽입합니다.
    cur.execute(
        """
        SELECT id FROM marks
        WHERE student_id = ? AND subject_id = ? AND exam = ?
        """,
        (student_id, subject_id, exam),
    )
    row = cur.fetchone()
    if row:
        cur.execute(
            "UPDATE marks SET marks_obtained = ? WHERE id = ?",
            (marks_obtained, row[0]),
        )
    else:
        cur.execute(
            """
            INSERT INTO marks (student_id, subject_id, exam, marks_obtained)
            VALUES (?, ?, ?, ?)
            """,
            (student_id, subject_id, exam, marks_obtained),
        )
    conn.commit()
    conn.close()


def get_marks_for_student(student_id, exam):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT s.code, s.name, s.max_marks, m.marks_obtained
        FROM subjects s
        LEFT JOIN marks m
            ON m.student_id = ?
            AND m.subject_id = s.id
            AND m.exam = ?
        ORDER BY s.code
        """,
        (student_id, exam),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


# --- 고급 분석 대화 상자 ---
class AdvancedWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("고급 기능")
        self.resize(900, 500)

        layout = QVBoxLayout()

        title = QLabel("고급 결과 분석")
        title.setStyleSheet("font-size: 16pt; font-weight: bold;")
        layout.addWidget(title)

        # 시험 입력
        exam_layout = QHBoxLayout()
        self.exam_input = QLineEdit()
        self.exam_input.setPlaceholderText("시험 입력 (e.g. 중간, 기말)")
        exam_layout.addWidget(QLabel("시험:"))
        exam_layout.addWidget(self.exam_input)
        layout.addLayout(exam_layout)

        # 버튼
        self.backlog_btn = QPushButton("학기 백로그 표시")
        self.rank_btn = QPushButton("학생 순위 표시(내림차순)")

        layout.addWidget(self.backlog_btn)
        layout.addWidget(self.rank_btn)

        # 표
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["학번", "이름", "합계 성적", "최고 점수", "비고"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)

        # ▼ 고급 분석 탭 폭 조정 ▼
        header = self.table.horizontalHeader()
        header.resizeSection(0, 100)  # 학번
        header.resizeSection(1, 200)  # 이름
        header.resizeSection(2, 100)  # 합계 성적
        header.resizeSection(3, 100)  # 최고 점수
        # ▲ 고급 분석 탭 폭 조정 ▲

        layout.addWidget(self.table)

        self.setLayout(layout)

        # 신호
        self.backlog_btn.clicked.connect(self.load_backlogs)
        self.rank_btn.clicked.connect(self.show_ranking)

    # ------------ 도우미: 통계 계산 -------------
    def compute_student_stats(self, exam):
        students = get_students()
        results = []

        for s_id, roll, name, cls, sec, dob in students:
            marks_rows = get_marks_for_student(s_id, exam)
            total = 0
            max_total = 0

            # 수정된 논리: 점수가 있는 경우에만 최대값을 계산합니다.
            for code, sname, maxm, mobt in marks_rows:
                if mobt is not None:
                    total += float(mobt)
                    max_total += float(maxm)
                # Note: Max marks for the subject is always counted if we want a true overall percentage
                # The implementation counts max_total ONLY when marks are obtained.
                # To be more accurate for ranking, max_total should be the sum of max_marks for all subjects
                # (or all subjects that the student is registered for).
                # We will stick to the provided logic (only counting max_total if marks exist) for consistency.
                elif mobt is None:
                    # If marks are missing, include the max marks in max_total for a more accurate 'possible' score
                    # For consistency with ranking based on obtained marks, we will stick to the original code's implied logic
                    # which tends to favor students with more recorded marks, but only if they are not None.
                    pass

            if max_total > 0:
                pct = (total / max_total) * 100
                results.append(
                    {
                        "student_id": s_id,
                        "roll": roll,
                        "name": name,
                        "total": total,
                        "max_total": max_total,
                        "percentage": pct,
                        "marks_rows": marks_rows,
                    }
                )
        return results

    # -------------------------- 1. Backlogs --------------------------
    def load_backlogs(self):
        exam = self.exam_input.text().strip()
        if not exam:
            QMessageBox.warning(self, "Error", "Enter exam name first!")
            return

        all_stats = self.compute_student_stats(exam)

        self.table.setRowCount(0)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["학번", "이름", "과목", "점수", "비고"]  # Roll No
        )

        row_i = 0
        PASS_PERCENT = 40.0

        for stu in all_stats:
            roll = stu["roll"]
            name = stu["name"]

            for code, sub_name, maxm, mobt in stu["marks_rows"]:
                if mobt is None:
                    continue

                percent = (float(mobt) / float(maxm)) * 100

                if percent < PASS_PERCENT:
                    self.table.insertRow(row_i)
                    self.table.setItem(row_i, 0, QTableWidgetItem(str(roll)))
                    self.table.setItem(row_i, 1, QTableWidgetItem(str(name)))
                    self.table.setItem(row_i, 2, QTableWidgetItem(str(sub_name)))
                    self.table.setItem(row_i, 3, QTableWidgetItem(str(mobt)))
                    self.table.setItem(row_i, 4, QTableWidgetItem("BACKLOG"))
                    row_i += 1

        if row_i == 0:
            QMessageBox.information(self, "Backlogs", "No backlogs found for this exam.")

    # ---------------------- 2. Ranking (descending = highest first) ----------------------
    def show_ranking(self):
        exam = self.exam_input.text().strip()
        if not exam:
            QMessageBox.warning(self, "Error", "Enter exam name first!")
            return

        all_stats = self.compute_student_stats(exam)

        if not all_stats:
            QMessageBox.information(self, "Ranking", "No marks found for this exam.")
            return

        # Corrected sorting – highest scorer FIRST
        all_stats.sort(key=lambda x: x["percentage"], reverse=True)

        self.table.setRowCount(0)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["학번", "이름", "합계 점수", "최고 점", "백분율"]
        )

        for i, s in enumerate(all_stats):
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(str(s["roll"])))
            self.table.setItem(i, 1, QTableWidgetItem(str(s["name"])))
            self.table.setItem(i, 2, QTableWidgetItem(f"{s['total']:.2f}"))
            self.table.setItem(i, 3, QTableWidgetItem(f"{s['max_total']:.2f}"))
            self.table.setItem(i, 4, QTableWidgetItem(f"{s['percentage']:.2f}%"))


# --- Stylesheet (Dark Blue Theme - for improved visibility) ---
DARK_BLUE_STYLESHEET = """
/* Global Settings */
QMainWindow {
    background-color: #f0f0f0; 
}

QWidget {
    background-color: #f0f0f0;
    color: #333333;
    font-size: 10pt;
}

/* Tabs */
QTabWidget::pane {
    border: 1px solid #0056b3; 
}

QTabBar::tab {
    background: #007bff; 
    color: #ffffff;      
    padding: 10px 20px;
    margin-right: 2px;
}

QTabBar::tab:selected {
    background: #0056b3; 
    font-weight: bold;
}

/* Input Fields */
QLineEdit, QComboBox, QSpinBox, QDateEdit {
    background-color: #ffffff; 
    border: 1px solid #cccccc; 
    padding: 5px;
    border-radius: 3px;
    color: #000000;
}

/* Buttons */
QPushButton {
    background-color: #007bff; 
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    color: #ffffff; 
    font-weight: bold;
}

QPushButton:hover {
    background-color: #0056b3; 
}

/* Table Widget */
QTableWidget {
    background-color: #ffffff; 
    gridline-color: #cccccc; 
    border: 1px solid #cccccc;
    selection-background-color: #cce5ff; 
    color: #000000;
}

QHeaderView::section {
    background-color: #dcdcdc; 
    padding: 6px;
    border: 1px solid #ffffff;
    color: #000000; 
    font-weight: bold;
}
"""


# --- Utility Functions ---
def show_error(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setWindowTitle("오류")
    msg.setText(str(message))
    msg.exec_()


def show_info(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setWindowTitle("정보")
    msg.setText(str(message))
    msg.exec_()


def grade_from_percentage(pct):
    if pct >= 90:
        return "A+"
    elif pct >= 80:
        return "A"
    elif pct >= 70:
        return "B+"
    elif pct >= 60:
        return "B"
    elif pct >= 50:
        return "C"
    elif pct >= 40:
        return "D"
    else:
        return "F"


# --- Student Management Tab ---
class StudentsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_student_id = None
        self.setup_ui()
        self.load_students()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()
        self.roll_input = QLineEdit()
        self.name_input = QLineEdit()
        self.class_input = QLineEdit()
        self.section_input = QLineEdit()
        self.dob_input = QDateEdit()
        self.dob_input.setCalendarPopup(True)
        self.dob_input.setDate(QDate.currentDate())

        form_layout.addRow("학번:", self.roll_input)
        form_layout.addRow("이름:", self.name_input)
        form_layout.addRow("학급:", self.class_input)
        form_layout.addRow("구분:", self.section_input)  # 학과
        form_layout.addRow("생년월일:", self.dob_input)

        layout.addLayout(form_layout)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("추가")
        self.update_btn = QPushButton("수정")
        self.delete_btn = QPushButton("삭제")
        self.clear_btn = QPushButton("Clear")

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.update_btn)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addWidget(self.clear_btn)

        layout.addLayout(btn_layout)

        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("학번, 이름 또는 학급으로 검색하세요")
        self.search_btn = QPushButton("검색")
        self.refresh_btn = QPushButton("새로고침")

        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_btn)
        search_layout.addWidget(self.refresh_btn)

        layout.addLayout(search_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["ID", "학번", "이름", "학급", "구분", "생일"]  # 구분 : 학과
        )
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setStretchLastSection(True)

        # ----------------------------------------------------
        # ▼ 학생 탭 테이블 폭 조정 ▼
        # ----------------------------------------------------
        header = self.table.horizontalHeader()
        header.resizeSection(0, 100)  # ID
        header.resizeSection(1, 180)  # 학번
        header.resizeSection(2, 280)  # 이름
        header.resizeSection(3, 110)  # 학급
        header.resizeSection(4, 160)  # 구분
        header.resizeSection(5, 360)  #
        # ----------------------------------------------------
        # ▲ 학생 탭 테이블 폭 조정 ▲
        # ----------------------------------------------------

        layout.addWidget(self.table)

        self.add_btn.clicked.connect(self.add_student)
        self.update_btn.clicked.connect(self.update_student)
        self.delete_btn.clicked.connect(self.delete_student)
        self.clear_btn.clicked.connect(self.clear_form)
        self.search_btn.clicked.connect(self.search_students)
        self.refresh_btn.clicked.connect(self.load_students)
        self.table.itemSelectionChanged.connect(self.table_selection_changed)

    def clear_form(self):
        self.selected_student_id = None
        self.roll_input.clear()
        self.name_input.clear()
        self.class_input.clear()
        self.section_input.clear()
        self.dob_input.setDate(QDate.currentDate())

        # 수정: itemSelectionChanged()가 트리거되지 않도록 테이블 선택을 지웁니다.
        self.table.blockSignals(True)
        self.table.clearSelection()
        self.table.blockSignals(False)

    def add_student(self):
        roll_no = self.roll_input.text().strip()
        name = self.name_input.text().strip()
        class_name = self.class_input.text().strip()
        section = self.section_input.text().strip()
        dob = self.dob_input.date().toString("yyyy-MM-dd")

        if not roll_no or not name or not class_name:
            show_error("학번, 이름, 학급이 필요합니다.")
            return

        try:
            add_student(roll_no, name, class_name, section, dob)
            show_info("학생 자료가 추가됬습니다.")
            self.load_students()
            self.clear_form()
        except Exception as e:
            show_error(f"학생을 추가하는 데 실패했습니다: {e}")

    def update_student(self):
        if not self.selected_student_id:
            show_error("표에서 수정할 학생을 선택하세요.")
            return

        roll_no = self.roll_input.text().strip()
        name = self.name_input.text().strip()
        class_name = self.class_input.text().strip()
        section = self.section_input.text().strip()
        dob = self.dob_input.date().toString("yyyy-MM-dd")

        if not roll_no or not name or not class_name:
            show_error("학번, 이름, 학급이 필요합니다.")
            return

        try:
            update_student(
                self.selected_student_id, roll_no, name, class_name, section, dob
            )
            show_info("학생 자료가 추가됬습니다.")
            self.load_students()
        except Exception as e:
            show_error(f"학생 수정에 실패했습니다: {e}")

    def delete_student(self):
        if not self.selected_student_id:
            show_error("표에서 삭제할 학생을 선택하세요.")
            return

        reply = QMessageBox.question(
            self,
            "삭제 확정",
            "이 학생과 관련된 모든 성적을 삭제하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.No:
            return

        try:
            delete_student(self.selected_student_id)
            show_info("학생 자료가 삭제되었습니다")
            self.load_students()
            self.clear_form()
        except Exception as e:
            show_error(f"학생 삭제에 실패했습니다: {e}")

    def load_students(self):
        try:
            students = get_students()
        except Exception as e:
            show_error(f"학생을 불러오는 데 실패했습니다: {e}")
            return

        self.table.setRowCount(0)
        for row_idx, row in enumerate(students):
            self.table.insertRow(row_idx)
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()

    def search_students(self):
        keyword = self.search_input.text().strip()
        try:
            if keyword:
                students = search_students(keyword)
            else:
                students = get_students()
        except Exception as e:
            show_error(f"학생 검색에 실패했습니다: {e}")
            return

        self.table.setRowCount(0)
        for row_idx, row in enumerate(students):
            self.table.insertRow(row_idx)
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()

    def table_selection_changed(self):
        selected = self.table.selectedItems()
        if not selected:
            self.selected_student_id = None
            return

        row = selected[0].row()
        self.selected_student_id = int(self.table.item(row, 0).text())
        self.roll_input.setText(self.table.item(row, 1).text())
        self.name_input.setText(self.table.item(row, 2).text())
        self.class_input.setText(self.table.item(row, 3).text())
        self.section_input.setText(self.table.item(row, 4).text())
        dob_text = self.table.item(row, 5).text()
        try:
            date = QDate.fromString(dob_text, "yyyy-MM-dd")
            if date.isValid():
                self.dob_input.setDate(date)
        except Exception:
            pass


# --- Subject Management Tab ---
class SubjectsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_subject_id = None
        self.setup_ui()
        self.load_subjects()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()
        self.code_input = QLineEdit()
        self.name_input = QLineEdit()
        self.max_marks_input = QSpinBox()
        self.max_marks_input.setRange(1, 1000)
        self.max_marks_input.setValue(100)

        form_layout.addRow("코드:", self.code_input)
        form_layout.addRow("이름:", self.name_input)
        form_layout.addRow("최고점:", self.max_marks_input)

        layout.addLayout(form_layout)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("추가")
        self.update_btn = QPushButton("수정")
        self.delete_btn = QPushButton("삭제")
        self.clear_btn = QPushButton("Clear")

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.update_btn)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addWidget(self.clear_btn)

        layout.addLayout(btn_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["ID", "코드", "이름", "최고점수"])
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setStretchLastSection(True)

        # ----------------------------------------------------
        # ▼ 과목 탭 테이블 폭 조정 ▼
        # ----------------------------------------------------
        header = self.table.horizontalHeader()
        header.resizeSection(0, 40)  # ID
        header.resizeSection(1, 80)  # 코드
        header.resizeSection(2, 200)  # 이름
        # ----------------------------------------------------
        # ▲ 과목 탭 테이블 폭 조정 ▲
        # ----------------------------------------------------

        layout.addWidget(self.table)

        self.add_btn.clicked.connect(self.add_subject)
        self.update_btn.clicked.connect(self.update_subject)
        self.delete_btn.clicked.connect(self.delete_subject)
        self.clear_btn.clicked.connect(self.clear_form)
        self.table.itemSelectionChanged.connect(self.table_selection_changed)

    def clear_form(self):
        self.selected_subject_id = None
        self.code_input.clear()
        self.name_input.clear()
        self.max_marks_input.setValue(100)

        self.table.blockSignals(True)
        self.table.clearSelection()
        self.table.blockSignals(False)

    def add_subject(self):
        code = self.code_input.text().strip()
        name = self.name_input.text().strip()
        max_marks = self.max_marks_input.value()

        if not code or not name:
            show_error("코드와 이름이 필요합니다.")
            return

        try:
            add_subject(code, name, max_marks)
            show_info("과목이 추가되었습니다.")
            self.load_subjects()
            self.clear_form()
        except Exception as e:
            show_error(f"과목을 추가하는 데 실패했습니다: {e}")

    def update_subject(self):
        if not self.selected_subject_id:
            show_error("표에서 수정할 과목을 선택하세요.")
            return

        code = self.code_input.text().strip()
        name = self.name_input.text().strip()
        max_marks = self.max_marks_input.value()

        if not code or not name:
            show_error("코드와 이름이 필요합니다.")
            return

        try:
            update_subject(self.selected_subject_id, code, name, max_marks)
            show_info("과목이 수정되었습니다.")
            self.load_subjects()
        except Exception as e:
            show_error(f"과목을 수정하지 못했습니다: {e}")

    def delete_subject(self):
        if not self.selected_subject_id:
            show_error("표에서 삭제할 과목을 선택하세요.")
            return

        reply = QMessageBox.question(
            self,
            "삭제 확정",
            "이 과목과 관련된 모든 점수를 삭제하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.No:
            return

        try:
            delete_subject(self.selected_subject_id)
            show_info("과목 삭제됨.")
            self.load_subjects()
            self.clear_form()
        except Exception as e:
            show_error(f"과목을 삭제하지 못했습니다: {e}")

    def load_subjects(self):
        try:
            subjects = get_subjects()
        except Exception as e:
            show_error(f"과목을 로드하는 데 실패했습니다: {e}")
            return

        self.table.setRowCount(0)
        for row_idx, row in enumerate(subjects):
            self.table.insertRow(row_idx)
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()

    def table_selection_changed(self):
        selected = self.table.selectedItems()
        if not selected:
            self.selected_subject_id = None
            return

        row = selected[0].row()
        self.selected_subject_id = int(self.table.item(row, 0).text())
        self.code_input.setText(self.table.item(row, 1).text())
        self.name_input.setText(self.table.item(row, 2).text())
        self.max_marks_input.setValue(int(self.table.item(row, 3).text()))


# --- 점수 입력 탭 ---
class MarksTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.students = []
        self.subjects = []
        self.setup_ui()
        self.load_students_subjects()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()
        self.student_combo = QComboBox()
        self.subject_combo = QComboBox()
        self.exam_input = QLineEdit()
        self.marks_input = QSpinBox()
        self.marks_input.setRange(0, 1000)

        form_layout.addRow("학생:", self.student_combo)
        form_layout.addRow("과목:", self.subject_combo)
        form_layout.addRow("시험 (e.g. 중간, 기말):", self.exam_input)
        form_layout.addRow("획득한 점수:", self.marks_input)

        layout.addLayout(form_layout)

        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton("점수 저장 / 수정")
        self.clear_btn = QPushButton("Clear")

        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.clear_btn)

        layout.addLayout(btn_layout)

        self.save_btn.clicked.connect(self.save_marks)
        self.clear_btn.clicked.connect(self.clear_form)

    def clear_form(self):
        if self.student_combo.count() > 0:
            self.student_combo.setCurrentIndex(0)

        if self.subject_combo.count() > 0:
            self.subject_combo.setCurrentIndex(0)

        self.exam_input.clear()
        self.marks_input.setValue(0)

    def load_students_subjects(self):
        try:
            self.students = get_students()
            self.subjects = get_subjects()
        except Exception as e:
            show_error(f"학생/과목을 불러오는 데 실패했습니다: {e}")
            return

        self.student_combo.clear()
        for s in self.students:
            _id, roll_no, name, class_name, section, dob = s
            self.student_combo.addItem(f"{roll_no} - {name} ({class_name})", _id)

        self.subject_combo.clear()
        for sub in self.subjects:
            _id, code, name, max_marks = sub
            self.subject_combo.addItem(f"{code} - {name}", _id)

    def save_marks(self):
        if self.student_combo.count() == 0 or self.subject_combo.count() == 0:
            show_error("먼저 학생과 과목을 추가하세요.")
            return

        student_id = self.student_combo.currentData()
        subject_id = self.subject_combo.currentData()
        exam = self.exam_input.text().strip()
        marks = self.marks_input.value()

        if not exam:
            show_error("시험명이 필요합니다")
            return

        # 선택한 과목에 대한 점수가 max_marks를 초과하지 않도록 합니다.
        selected_sub_index = self.subject_combo.currentIndex()
        if selected_sub_index >= 0:
            sub = self.subjects[selected_sub_index]
            max_marks = sub[3]  # sub[3] is max_marks
            if marks > max_marks:
                show_error(f"획득한 점수({marks})는 최대 점수({max_marks})를 초과할 수 없습니다.")
                return

        try:
            set_marks(student_id, subject_id, exam, marks)
            show_info("점수가 저장되었습니다.")
        except Exception as e:
            show_error(f"점수를 저장하지 못했습니다: {e}")


# --- Reports and Export Tab ---
class ReportsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_student = None
        self.current_marks_data = None
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        top_layout = QHBoxLayout()
        self.roll_input = QLineEdit()
        self.roll_input.setPlaceholderText("학번 입력")
        self.exam_input = QLineEdit()
        self.exam_input.setPlaceholderText("시험 (e.g. 중간, 기말)")
        self.load_btn = QPushButton("결과 로드")
        self.export_btn = QPushButton("CSV 매보내기")

        top_layout.addWidget(QLabel("학번:"))
        top_layout.addWidget(self.roll_input)
        top_layout.addWidget(QLabel("시험:"))
        top_layout.addWidget(self.exam_input)
        top_layout.addWidget(self.load_btn)
        top_layout.addWidget(self.export_btn)

        layout.addLayout(top_layout)

        self.info_label = QLabel("결과 세부정보가 여기에 표시됩니다.")
        layout.addWidget(self.info_label)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(
            ["과목 코드", "과목 이름", "최고 점수", "획득 점수"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)

        # ----------------------------------------------------
        # ▼ 보고서 탭 테이블 폭 조정 ▼
        # ----------------------------------------------------
        header = self.table.horizontalHeader()
        header.resizeSection(0, 100)  # 과목 코드
        header.resizeSection(1, 250)  # 과목 이름
        header.resizeSection(2, 100)  # 최고 점수
        # ----------------------------------------------------
        # ▲ 보고서 탭 테이블 폭 조정 ▲
        # ----------------------------------------------------

        layout.addWidget(self.table)

        self.summary_label = QLabel("")
        layout.addWidget(self.summary_label)

        self.load_btn.clicked.connect(self.load_result)
        self.export_btn.clicked.connect(self.export_csv)

    def load_result(self):
        roll_no = self.roll_input.text().strip()
        exam = self.exam_input.text().strip()

        if not roll_no or not exam:
            show_error("학번과 시험명을 입력하세요.")
            return

        try:
            student = get_student_by_roll(roll_no)
        except Exception as e:
            show_error(f"학생을 불러오는 데 실패했습니다: {e}")
            return

        if not student:
            show_error("해당 학번을 가진 학생이 없습니다.")
            return

        self.current_student = student
        student_id, roll_no, name, class_name, section, dob = student

        try:
            marks_rows = get_marks_for_student(student_id, exam)
        except Exception as e:
            show_error(f"점수를 로드하는 데 실패했습니다: {e}")
            return

        self.current_marks_data = marks_rows

        self.info_label.setText(
            f"Name: {name} | Roll: {roll_no} | Class: {class_name} | Section: {section} | Exam: {exam}"
        )

        self.table.setRowCount(0)
        total_obt = 0
        total_max = 0
        has_marks = False

        for row_idx, (code, sub_name, max_marks, marks_obt) in enumerate(marks_rows):
            self.table.insertRow(row_idx)
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(code)))
            self.table.setItem(row_idx, 1, QTableWidgetItem(str(sub_name)))
            self.table.setItem(row_idx, 2, QTableWidgetItem(str(max_marks)))

            marks_text = str(marks_obt) if marks_obt is not None else "N/A"
            self.table.setItem(row_idx, 3, QTableWidgetItem(marks_text))

            if marks_obt is not None:
                total_obt += float(marks_obt)
                total_max += float(max_marks)
                has_marks = True

        percentage = (total_obt / total_max) * 100 if total_max > 0 else 0
        grade = grade_from_percentage(percentage)

        if has_marks:
            summary = (
                f"총점: {total_obt:.2f} / {total_max:.2f} | "
                f"백분율: {percentage:.2f}% | "
                f"최종 등급: {grade}"
            )
        else:
            summary = "이 시험에 대한 성적을 찾을 수 없습니다."

        self.summary_label.setText(summary)

    def export_csv(self):
        if not self.current_marks_data:
            show_error("먼저 결과를 로드하세요.")
            return

        # 파일 저장 대화 상자를 엽니다.
        path, _ = QFileDialog.getSaveFileName(
            self, "CSV 파일 저장", "", "CSV Files (*.csv)"
        )

        if path:
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)

                    # 1. 헤더 (학생 정보)
                    writer.writerow(["학번", "이름", "학급", "구분", "시험"])
                    writer.writerow([
                        self.current_student[1],
                        self.current_student[2],
                        self.current_student[3],
                        self.current_student[4],
                        self.exam_input.text().strip()
                    ])
                    writer.writerow([])

                    # 2. 과목/성적 헤더
                    writer.writerow(["과목 코드", "과목 이름", "최고 점수", "획득 점수"])

                    # 3. 과목/성적 데이터
                    total_obt = 0
                    total_max = 0
                    for row in self.current_marks_data:
                        writer.writerow([str(r) if r is not None else "N/A" for r in row])
                        if row[3] is not None:
                            total_obt += float(row[3])
                            total_max += float(row[2])

                    # 4. 요약
                    percentage = (total_obt / total_max) * 100 if total_max > 0 else 0
                    grade = grade_from_percentage(percentage)

                    writer.writerow([])
                    writer.writerow(["총점", f"{total_obt:.2f} / {total_max:.2f}"])
                    writer.writerow(["백분율", f"{percentage:.2f}%"])
                    writer.writerow(["최종 등급", grade])

                show_info(f"결과가 성공적으로 저장되었습니다: {path}")
            except Exception as e:
                show_error(f"CSV 파일 저장에 실패했습니다: {e}")


# --- 메인 윈도우 ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("학생 성적 관리 Ver 1.0_251128")
        self.setGeometry(100, 100, 900, 700)
        self.setWindowIcon(QIcon("icon.png"))  # 아이콘 파일을 준비해야 합니다.

        # 스타일 시트 적용
        self.setStyleSheet(DARK_BLUE_STYLESHEET)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.students_tab = StudentsTab()
        self.subjects_tab = SubjectsTab()
        self.marks_tab = MarksTab()
        self.reports_tab = ReportsTab()

        self.tabs.addTab(self.students_tab, "1. 학생 관리")
        self.tabs.addTab(self.subjects_tab, "2. 과목 관리")
        self.tabs.addTab(self.marks_tab, "3. 성적 입력")
        self.tabs.addTab(self.reports_tab, "4. 보고서/내보내기")

        # 고급 분석 버튼
        adv_btn = QPushButton("고급 분석 (순위, 백로그)")
        adv_btn.clicked.connect(self.show_advanced_window)

        # 탭 아래에 버튼 추가
        central_layout = QVBoxLayout()
        central_layout.addWidget(self.tabs)
        central_layout.addWidget(adv_btn)

        central_widget = QWidget()
        central_widget.setLayout(central_layout)
        self.setCentralWidget(central_widget)

    def show_advanced_window(self):
        adv_window = AdvancedWindow()
        adv_window.exec_()

    def closeEvent(self, event):
        # 애플리케이션 종료 시 데이터베이스 연결이 안전하게 종료되도록 보장합니다.
        event.accept()


# --- 실행 ---
if __name__ == "__main__":
    init_db()  # DB 초기화

    app = QApplication(sys.argv)

    # 윈도우 스타일 설정은 QMainWindow에서 처리했습니다.
    # app.setStyleSheet(DARK_BLUE_STYLESHEET)

    main_win = MainWindow()
    main_win.show()

    sys.exit(app.exec_())