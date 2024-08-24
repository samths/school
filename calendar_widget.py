"""
calendar_widget.py   달력 위젯  Ver 1.0_240825
"""
import sys
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget, QVBoxLayout, QCalendarWidget, QLabel
from PyQt5.QtGui import QIcon, QTextCharFormat, QColor, QFont
import warnings

warnings.filterwarnings("ignore", category=DeprecationWarning)


class MyApp(QWidget):
    # 초기화 동작 수행
    def __init__(self):
        super().__init__()
        self.initUI()  # 기본 UI 초기화

    # 윈도우를 화면 가운데 위치 시키기
    def center(self):
        qr = self.frameGeometry()  # 창의 위치와 크기 정보를 가져옴
        cp = QDesktopWidget().availableGeometry().center()  # 사용하는 모니터 화면의 가운데 위치를 파악
        qr.moveCenter(cp)  # 창의 직사각형 위치를 화면의 중심의 위치로 이동합니다.
        self.move(qr.topLeft())

    # 화면 기본 설정
    def initUI(self):
        # 달력
        cal = QCalendarWidget(self)
        cal.setGridVisible(True)

        # 토요일의 서식 변경 (파란색)
        saturday_format = QTextCharFormat()
        saturday_format.setForeground(QColor('blue'))
        cal.setWeekdayTextFormat(Qt.Saturday, saturday_format)

        # 현재 달의 모든 날짜를 진하게 설정
        bold_format = QTextCharFormat()
        bold_format.setFontWeight(QFont.Bold)  # 폰트를 진하게 설정

        # 현재 달의 첫 번째와 마지막 날짜 계산
        current_date = QDate.currentDate()
        first_day = QDate(current_date.year(), current_date.month(), 1)
        last_day = QDate(current_date.year(), current_date.month(), current_date.daysInMonth())

        # 모든 날짜에 진하게 설정 적용
        for day in range(1, current_date.daysInMonth() + 1):
            date = QDate(current_date.year(), current_date.month(), day)
            cal.setDateTextFormat(date, bold_format)

        # 날짜를 클릭했을 때 showDate 메서드가 호출
        cal.clicked[QDate].connect(self.showDate)

        self.lbl = QLabel(self)
        date = cal.selectedDate()
        self.lbl.setText(date.toString(Qt.DefaultLocaleLongDate))

        # 레이아웃 설정
        vBox = QVBoxLayout()
        vBox.addWidget(cal)
        vBox.addWidget(self.lbl)

        self.setLayout(vBox)

        self.setWindowTitle("달력 위젯")  # 타이틀
        self.setWindowIcon(QIcon('../../Utils/Images/web.png'))  # 아이콘 추가
        self.resize(500, 350)  # 크기 설정
        self.center()
        self.show()  # 보이기

    def showDate(self, date):
        self.lbl.setText(date.toString(Qt.DefaultLocaleLongDate))


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 애플리케이션의 기본 폰트를 '맑은 고딕'으로 설정
    font = QFont('맑은 고딕', 12)
    app.setFont(font)

    ex = MyApp()
    sys.exit(app.exec())
