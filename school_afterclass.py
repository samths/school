"""
school_afterclass.py 학교 방과후 수럽 출석부 작성 Ver 1.0_250116
"""
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, QVBoxLayout, QWidget, QMessageBox
)
import pandas as pd
from pyhwpx import Hwp

class AttendanceApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.xlsx_path = ""
        self.hwpx_path = ""

        self.initUI()

    def initUI(self):
        self.setWindowTitle("출석부 생성 프로그램")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        self.xlsx_label = QLabel("학생명단 엑셀파일: 선택되지 않음", self)
        layout.addWidget(self.xlsx_label)

        self.hwpx_label = QLabel("출석부 서식 한글파일: 선택되지 않음", self)
        layout.addWidget(self.hwpx_label)

        xlsx_button = QPushButton("엑셀 파일 선택", self)
        xlsx_button.clicked.connect(self.select_xlsx)
        layout.addWidget(xlsx_button)

        hwpx_button = QPushButton("한글 파일 선택", self)
        hwpx_button.clicked.connect(self.select_hwpx)
        layout.addWidget(hwpx_button)

        generate_button = QPushButton("출석부 생성", self)
        generate_button.clicked.connect(self.generate_attendance)
        layout.addWidget(generate_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_xlsx(self):
        # 명단 자료 : 학생명단.xlsx
        file_path, _ = QFileDialog.getOpenFileName(self, "학생명단 엑셀파일을 선택해주세요.", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.xlsx_path = file_path
            self.xlsx_label.setText(f"학생명단 엑셀파일: {file_path}")

    def select_hwpx(self):
        # 출석부 서식 : 방과후 프로그램 출석부 서식.hwpx
        file_path, _ = QFileDialog.getOpenFileName(self, "출석부 서식 한글파일을 선택해주세요.", "", "HWPX Files (*.hwpx)")
        if file_path:
            self.hwpx_path = file_path
            self.hwpx_label.setText(f"출석부 서식 한글파일: {file_path}")

    def generate_attendance(self):
        if not self.xlsx_path or not self.hwpx_path:
            QMessageBox.warning(self, "경고", "모든 파일을 선택해주세요.")
            return

        try:
            df = pd.read_excel(self.xlsx_path)
            hwp = Hwp(visible=False)

            for cat in df["프로그램명"].unique():
                hwp.open(self.hwpx_path)
                df_cat = df[df["프로그램명"] == cat]
                hwp.put_field_text(df_cat)
                hwp.put_field_text("연번", list(range(1, len(df_cat) + 1)))
                hwp.save_as(f"./school/방과후프로그램출석부_{cat}.hwpx")
                hwp.clear()

            hwp.msgbox("작업이 완료되었습니다.")
            hwp.quit()
            QMessageBox.information(self, "완료", "출석부 생성이 완료되었습니다.")

        except Exception as e:
            QMessageBox.critical(self, "오류", f"오류가 발생했습니다: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = AttendanceApp()
    main_window.show()
    sys.exit(app.exec_())
