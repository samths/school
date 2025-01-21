"""
book_search.py   도서 검색 관리 Ver 1.0_250111
"""
import sys
import sqlite3
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QTabWidget,
                             QHBoxLayout, QLineEdit, QPushButton, QLabel, QListWidget,
                             QTextEdit, QMessageBox, QComboBox, QDialog, QFormLayout,
                             QAction, QSizePolicy)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt


STYLE_SHEET = """
QMainWindow {
    background-color: rgb(240, 240, 240);
}

QTabWidget::pane {
    border: 1px solid #cccccc;
    background: white;
    border-radius: 4px;
}

QTabWidget::tab-bar {
    left: 5px;
}

QTabBar::tab {
    background: #e1e1e1;
    border: 1px solid #cccccc;
    padding: 8px 12px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    font-family: "맑은 고딕";
}

QTabBar::tab:selected {
    background: white;
    border-bottom-color: white;
}

QLineEdit, QComboBox {
    padding: 8px;
    border: 1px solid #cccccc;
    border-radius: 4px;
    background-color: white;
    font-family: "맑은 고딕";
}

QLineEdit:focus, QComboBox:focus {
    border: 1px solid #4a90e2;
}

QPushButton {
    background-color: #4a90e2;
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    min-width: 100px;
    font-family: "맑은 고딕";
}

QPushButton:hover {
    background-color: #357abd;
}

QPushButton:pressed {
    background-color: #2d6da3;
}

QListWidget {
    border: 1px solid #cccccc;
    border-radius: 4px;
    background-color: white;
    padding: 5px;
    font-family: "맑은 고딕";
}

QListWidget::item {
    padding: 8px;
    border-bottom: 1px solid #eeeeee;
    font-family: "맑은 고딕";
}

QListWidget::item:selected {
    background-color: #4a90e2;
    color: white;
}

QTextEdit {
    border: 1px solid #cccccc;
    border-radius: 4px;
    background-color: white;
    padding: 5px;
    font-family: "맑은 고딕";
}

QLabel {
    color: #333333;
    font-family: "맑은 고딕";
    font-size: 12pt;
}
"""

class AddEditBookDialog(QDialog):
    def __init__(self, parent=None, book_data=None):
        super().__init__(parent)
        self.setWindowTitle("책 추가/수정")
        self.setModal(True)
        self.book_data = book_data
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout(self)

        self.title_input = QLineEdit()
        self.author_input = QLineEdit()
        self.genre_combo = QComboBox()
        self.genre_combo.addItems(["철학", "종교", "사회과학", "자연과학", "기술과학", "예술", "언어", "문학", "역사"])
        self.description_input = QTextEdit()

        if self.book_data:
            self.title_input.setText(self.book_data['title'])
            self.author_input.setText(self.book_data['author'])
            self.genre_combo.setCurrentText(self.book_data['genre'])
            self.description_input.setText(self.book_data['description'])
            self.title_input.setReadOnly(True)

        layout.addRow("제목:", self.title_input)
        layout.addRow("저자:", self.author_input)
        layout.addRow("장르:", self.genre_combo)
        layout.addRow("설명:", self.description_input)

        button_layout = QHBoxLayout()
        save_btn = QPushButton("저장")
        cancel_btn = QPushButton("취소")

        save_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow("", button_layout)

    def get_book_data(self):
        return {
            'title': self.title_input.text(),
            'author': self.author_input.text(),
            'genre': self.genre_combo.currentText(),
            'description': self.description_input.toPlainText()
        }


class BookManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("도서 관리 시스템")
        self.setGeometry(600, 200, 800, 600)
        self.setWindowIcon(QIcon("book_picture.png"))
        self.setStyleSheet(STYLE_SHEET)

        self.init_database()

        app_font = QFont("맑은 고딕", 10)
        QApplication.setFont(app_font)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.setup_search_tab()
        self.setup_details_tab()
        self.setup_management_tab()

        layout.addWidget(self.tabs)

    def init_database(self):
        self.conn = sqlite3.connect('./data/books.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                title TEXT PRIMARY KEY,
                author TEXT NOT NULL,
                genre TEXT NOT NULL,
                description TEXT
            )
        ''')
        self.conn.commit()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("도서 관리 시스템")
        self.setGeometry(600, 200, 800, 600)
        self.setWindowIcon(QIcon("book_picture.png"))
        self.setStyleSheet(STYLE_SHEET)

        self.init_database()

        app_font = QFont("맑은 고딕", 10)
        QApplication.setFont(app_font)

        # 메뉴바 설정
        self.setup_menu_bar()

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.setup_search_tab()
        self.setup_details_tab()
        self.setup_management_tab()

        layout.addWidget(self.tabs)


    def setup_menu_bar(self):
        """메뉴바에 Help 메뉴를 추가하고 오른쪽 끝에 배치."""
        menu_bar = self.menuBar()

        # 빈 위젯 추가로 오른쪽 정렬 구현
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        menu_bar.setCornerWidget(spacer, Qt.TopRightCorner)

        # Help 메뉴 추가
        help_menu = menu_bar.addMenu("Help")

        # 작성자 정보 메뉴
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        # 사용법 메뉴
        usage_action = QAction("Usage", self)
        usage_action.triggered.connect(self.show_usage_dialog)
        help_menu.addAction(usage_action)

    def show_about_dialog(self):
        """작성자 정보를 표시."""
        QMessageBox.information(
            self, "About",
            "도서 관리 시스템 Ver 1.0\n\n"
            "작성자: 윤용수\n"
            "Email: samths@naver.com\n"
            "개발 날짜: 2025-01-11\n"
        )

    def show_usage_dialog(self):
        """프로그램 사용법을 표시."""
        QMessageBox.information(
            self, "사용법",
            "1. 도서 검색:\n"
            "   - 상단 검색 탭에서 책 제목 또는 저자를 입력해 검색.\n"
            "2. 도서 추가:\n"
            "   - 도서 관리 탭에서 '새 책 추가' 버튼을 눌러 새 책을 추가\n"
            "3. 도서 수정 및 삭제:\n"
            "   - 도서 관리 탭에서 도서를 선택 후 '수정' 또는 '삭제'\n"
            "     버튼을 사용\n"
            "4. 상세 정보:\n"
            "   - 검색 또는 도서 관리 탭에서 항목을 더블 클릭하면 \n"
            "     상세 정보를 볼 수 있습니다."
        )

    #  탭 검색
    def setup_search_tab(self):
        search_tab = QWidget()
        layout = QVBoxLayout()

        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("책 제목 입력...")
        self.search_input.returnPressed.connect(self.search_books)

        self.genre_filter = QComboBox()
        self.genre_filter.addItems(["전체", "철학", "종교", "사회과학", "자연과학", "기술과학", "예술", "언어", "문학", "역사"])

        search_button = QPushButton("검색")
        search_button.clicked.connect(self.search_books)

        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.genre_filter)
        filter_layout.addWidget(search_button)
        layout.addLayout(filter_layout)

        results_label = QLabel("검색 결과:")
        layout.addWidget(results_label)

        self.results_list = QListWidget()
        self.results_list.itemDoubleClicked.connect(self.view_book_details)
        layout.addWidget(self.results_list)

        view_details_btn = QPushButton("상세 정보")
        view_details_btn.clicked.connect(self.view_book_details)
        layout.addWidget(view_details_btn)

        search_tab.setLayout(layout)
        self.tabs.addTab(search_tab, "도서 검색")

    # 세부 사항
    def setup_details_tab(self):
        details_tab = QWidget()
        layout = QVBoxLayout()

        self.book_title_label = QLabel("제목: ")
        self.book_author_label = QLabel("저자: ")
        self.book_genre_label = QLabel("장르: ")
        self.book_description = QTextEdit()
        self.book_description.setReadOnly(True)

        layout.addWidget(self.book_title_label)
        layout.addWidget(self.book_author_label)
        layout.addWidget(self.book_genre_label)
        layout.addWidget(QLabel("설명:"))
        layout.addWidget(self.book_description)

        button_layout = QHBoxLayout()
        edit_btn = QPushButton("수정")
        edit_btn.clicked.connect(self.edit_book)
        delete_btn = QPushButton("삭제")
        delete_btn.clicked.connect(self.delete_book)
        back_btn = QPushButton("돌아가기")
        back_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(0))

        button_layout.addWidget(back_btn)
        button_layout.addWidget(edit_btn)
        button_layout.addWidget(delete_btn)
        layout.addLayout(button_layout)

        details_tab.setLayout(layout)
        self.tabs.addTab(details_tab, "상세 정보")

    # 탭 설정
    def setup_management_tab(self):
        management_tab = QWidget()
        layout = QVBoxLayout()

        # 새 책 추가 버튼
        add_book_btn = QPushButton("새 책 추가")
        add_book_btn.clicked.connect(self.add_new_book)
        layout.addWidget(add_book_btn)  # "도서 목록" 위에 버튼 추가

        layout.addWidget(QLabel("전체 도서 목록:"))
        self.books_list = QListWidget()
        self.books_list.itemClicked.connect(self.update_buttons_state)
        layout.addWidget(self.books_list)

        button_layout = QHBoxLayout()

        refresh_btn = QPushButton("목록 새로고침")
        refresh_btn.clicked.connect(self.refresh_books_list)
        button_layout.addWidget(refresh_btn)

        edit_btn = QPushButton("수정")
        edit_btn.setEnabled(False)
        edit_btn.clicked.connect(self.edit_selected_book)
        button_layout.addWidget(edit_btn)

        delete_btn = QPushButton("삭제")
        delete_btn.setEnabled(False)
        delete_btn.clicked.connect(self.delete_selected_book)
        button_layout.addWidget(delete_btn)

        layout.addLayout(button_layout)

        management_tab.setLayout(layout)
        self.tabs.addTab(management_tab, "도서 관리")

        self.refresh_books_list()
        self.edit_btn = edit_btn
        self.delete_btn = delete_btn

    # 도서 검색
    def search_books(self):
        search_term = self.search_input.text()
        genre_filter = self.genre_filter.currentText()

        self.results_list.clear()

        try:
            if genre_filter == "전체":
                self.cursor.execute('''
                    SELECT title, author, genre FROM books 
                    WHERE title LIKE ? OR author LIKE ?
                ''', (f'%{search_term}%', f'%{search_term}%'))
            else:
                self.cursor.execute('''
                    SELECT title, author, genre FROM books 
                    WHERE (title LIKE ? OR author LIKE ?) AND genre = ?
                ''', (f'%{search_term}%', f'%{search_term}%', genre_filter))

            results = self.cursor.fetchall()

            for title, author, genre in results:
                self.results_list.addItem(f"{title} - {author} ({genre})")

            if self.results_list.count() == 0:
                self.results_list.addItem("검색 결과가 없습니다.")

        except sqlite3.Error as e:
            QMessageBox.critical(self, "데이터베이스 오류", f"검색 중 오류가 발생했습니다: {str(e)}")

    # 도서 나열
    def view_book_details(self):
        """선택한 책의 상세 정보를 표시합니다."""
        selected_item = self.results_list.currentItem() or self.books_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "책을 선택해주세요.")
            return

        title = selected_item.text().split(" - ")[0]

        try:
            self.cursor.execute('SELECT * FROM books WHERE title = ?', (title,))
            book = self.cursor.fetchone()

            if book:
                self.book_title_label.setText(f"제목: {book[0]}")
                self.book_author_label.setText(f"저자: {book[1]}")
                self.book_genre_label.setText(f"장르: {book[2]}")
                self.book_description.setText(book[3])
                self.tabs.setCurrentIndex(1)  # 상세 정보 탭으로 이동
            else:
                QMessageBox.warning(self, "오류", "선택한 책의 정보를 찾을 수 없습니다.")

        except sqlite3.Error as e:
            QMessageBox.critical(self, "데이터베이스 오류", f"상세 정보 로드 중 오류가 발생했습니다: {str(e)}")

    # 도서 수정
    def edit_book(self):
        """상세 정보 탭에서 현재 선택된 책을 수정."""
        current_title = self.book_title_label.text().replace("제목: ", "")
        if not current_title:
            QMessageBox.warning(self, "선택 오류", "수정할 책이 선택되지 않았습니다.")
            return

        try:
            self.cursor.execute('SELECT * FROM books WHERE title = ?', (current_title,))
            book = self.cursor.fetchone()
            if book:
                book_data = {
                    'title': book[0],
                    'author': book[1],
                    'genre': book[2],
                    'description': book[3]
                }

                dialog = AddEditBookDialog(self, book_data)
                if dialog.exec_() == QDialog.Accepted:
                    updated_data = dialog.get_book_data()
                    self.cursor.execute('''
                        UPDATE books 
                        SET author = ?, genre = ?, description = ?
                        WHERE title = ?
                    ''', (updated_data['author'], updated_data['genre'],
                          updated_data['description'], current_title))
                    self.conn.commit()
                    self.refresh_books_list()
                    self.view_book_details()  # 상세 정보 새로고침
                    QMessageBox.information(self, "성공", "책 정보가 수정되었습니다.")
            else:
                QMessageBox.warning(self, "오류", "선택한 책 정보를 찾을 수 없습니다.")

        except sqlite3.Error as e:
            QMessageBox.critical(self, "데이터베이스 오류", f"책 수정 중 오류가 발생했습니다: {str(e)}")

    # 도서 추가
    def add_new_book(self):
        """새로운 책을 추가하기 위한 다이얼로그를 표시합니다."""
        dialog = AddEditBookDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            book_data = dialog.get_book_data()
            try:
                self.cursor.execute('''
                    INSERT INTO books (title, author, genre, description)
                    VALUES (?, ?, ?, ?)
                ''', (book_data['title'], book_data['author'],
                      book_data['genre'], book_data['description']))
                self.conn.commit()
                self.refresh_books_list()
                QMessageBox.information(self, "성공", "새 책이 추가되었습니다.")
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "중복 오류", "같은 제목의 책이 이미 존재합니다.")
            except sqlite3.Error as e:
                QMessageBox.critical(self, "데이터베이스 오류", f"책 추가 중 오류가 발생했습니다: {str(e)}")

    # 수정 버튼 상태
    def update_buttons_state(self):
        selected_item = self.books_list.currentItem()
        self.edit_btn.setEnabled(bool(selected_item))
        self.delete_btn.setEnabled(bool(selected_item))

    # 도서 나열
    def refresh_books_list(self):
        self.books_list.clear()
        try:
            self.cursor.execute('SELECT title, author, genre FROM books ORDER BY title')
            books = self.cursor.fetchall()
            for title, author, genre in books:
                self.books_list.addItem(f"{title} - {author} ({genre})")
        except sqlite3.Error as e:
            QMessageBox.critical(self, "데이터베이스 오류", f"도서 목록 새로고침 중 오류가 발생했습니다: {str(e)}")

    # 선택 도서 수정
    def edit_selected_book(self):
        selected_item = self.books_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "수정할 책을 선택해주세요.")
            return

        title = selected_item.text().split(" - ")[0]
        try:
            self.cursor.execute('SELECT * FROM books WHERE title = ?', (title,))
            book = self.cursor.fetchone()
            if book:
                book_data = {
                    'title': book[0],
                    'author': book[1],
                    'genre': book[2],
                    'description': book[3]
                }

                dialog = AddEditBookDialog(self, book_data)
                if dialog.exec_() == QDialog.Accepted:
                    updated_data = dialog.get_book_data()
                    self.cursor.execute('''
                        UPDATE books 
                        SET author = ?, genre = ?, description = ?
                        WHERE title = ?
                    ''', (updated_data['author'], updated_data['genre'],
                          updated_data['description'], title))
                    self.conn.commit()
                    self.refresh_books_list()
                    QMessageBox.information(self, "성공", "책 정보가 수정되었습니다.")
        except sqlite3.Error as e:
            QMessageBox.critical(self, "데이터베이스 오류", f"책 수정 중 오류가 발생했습니다: {str(e)}")

    # 도서 삭제
    def delete_book(self):
        """상세 정보 탭에서 현재 선택된 책을 삭제."""
        current_title = self.book_title_label.text().replace("제목: ", "")
        if not current_title:
            QMessageBox.warning(self, "선택 오류", "삭제할 책이 선택되지 않았습니다.")
            return

        reply = QMessageBox.question(self, "삭제 확인",
                                     f"'{current_title}'을(를) 삭제하시겠습니까?",
                                     QMessageBox.Yes | QMessageBox.No)

        if reply == QMessageBox.Yes:
            try:
                self.cursor.execute('DELETE FROM books WHERE title = ?', (current_title,))
                self.conn.commit()
                self.refresh_books_list()
                self.tabs.setCurrentIndex(0)  # 검색 탭으로 이동
                QMessageBox.information(self, "성공", "책이 삭제되었습니다.")
            except sqlite3.Error as e:
                QMessageBox.critical(self, "데이터베이스 오류", f"책 삭제 중 오류가 발생했습니다: {str(e)}")

    # 선택 도서 삭제
    def delete_selected_book(self):
        selected_item = self.books_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "삭제할 책을 선택해주세요.")
            return

        title = selected_item.text().split(" - ")[0]
        reply = QMessageBox.question(self, "삭제 확인",
                                     f"'{title}'을(를) 삭제하시겠습니까?",
                                     QMessageBox.Yes | QMessageBox.No)

        if reply == QMessageBox.Yes:
            try:
                self.cursor.execute('DELETE FROM books WHERE title = ?', (title,))
                self.conn.commit()
                self.refresh_books_list()
                QMessageBox.information(self, "성공", "책이 삭제되었습니다.")
            except sqlite3.Error as e:
                QMessageBox.critical(self, "데이터베이스 오류", f"책 삭제 중 오류가 발생했습니다: {str(e)}")

    # 종료
    def closeEvent(self, event):
        if hasattr(self, 'conn'):
            self.conn.close()
        event.accept()

# 메인 프로그렘
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = BookManagerApp()
    window.show()
    sys.exit(app.exec_())
