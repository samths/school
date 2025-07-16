"""
xlsx3sqlite3.py   xlsx to sqlit3 converter   Ver 1.0_250716
"""
import sys
from PyQt6 import QtWidgets, uic
from PyQt6.QtWidgets import QFileDialog, QMessageBox
import pandas as pd
import sqlite3

class XlsxToSQLiteApp(QtWidgets.QDialog):
    def __init__(self):
        super(XlsxToSQLiteApp, self).__init__()

        # UI를 동적으로 로드
        self.setWindowTitle("Excel to SQLite 변환기")
        self.setGeometry(300, 300, 400, 200)

        # UI 요소 만들기
        self.loadFileButton = QtWidgets.QPushButton("Excel 파일 로드", self)
        self.loadFileButton.setGeometry(50, 50, 300, 40)
        self.loadFileButton.clicked.connect(self.load_file)

        self.saveToDBButton = QtWidgets.QPushButton("SQLite에 저장", self)
        self.saveToDBButton.setGeometry(50, 100, 300, 40)
        self.saveToDBButton.clicked.connect(self.save_to_db)
        self.saveToDBButton.setEnabled(False)

        self.statusLabel = QtWidgets.QLabel("상태: 입력 대기 중...", self)
        self.statusLabel.setGeometry(50, 150, 300, 30)

        # 변수 초기화
        self.excel_data = None
        self.file_path = None

    def load_file(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(self, 'Open Excel File', '', 'Excel files (*.xlsx)')
            if file_path:
                self.excel_data = pd.read_excel(file_path, sheet_name=None)
                self.file_path = file_path
                self.statusLabel.setText("Excel 파일이 성공적으로 로드되었습니다.")
                QMessageBox.information(self, "성공", "Excel 파일이 성공적으로 로드되었습니다.")
                self.saveToDBButton.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "오류", f"{e} 파일을 로드하는 데 실패했습니다.")

    def save_to_db(self):
        if not self.file_path:
            QMessageBox.warning(self, "경고", "파일이 로드되지 않았습니다.")
            return

        db_path, _ = QFileDialog.getSaveFileName(self, 'Save Database', '', 'SQLite Database (*.db)')
        if db_path:
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()

                # 각 시트를 데이터베이스에 저장
                for sheet_name, df in self.excel_data.items():
                    cursor.execute(f"DROP TABLE IF EXISTS '{sheet_name}'")
                    columns_definitions = ", ".join([f'"{col}" TEXT' for col in df.columns])
                    sql_create_table = f"CREATE TABLE '{sheet_name}' ({columns_definitions});"
                    cursor.execute(sql_create_table)
                    df.to_sql(sheet_name, con=cursor.connection, if_exists='append', index=False)

                conn.commit()
                conn.close()
                self.statusLabel.setText("데이터베이스가 성공적으로 저장되었습니다.")
                QMessageBox.information(self, "성공", "데이터베이스가 성공적으로 저장되었습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"{e} : 데이터베이스 저장에 실패했습니다")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    dialog = XlsxToSQLiteApp()
    dialog.show()
    sys.exit(app.exec())