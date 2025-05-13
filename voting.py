"""
voting.py   투표 프로그램   Ver 1.0_250514
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import sqlite3 as sqltor
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = "Malgun Gothic"  # 그래프에서 한글 글꼴 깨짐 방지
plt.rcParams['axes.unicode_minus'] = False

# 메인 데이터베이스 연결
conn = sqltor.connect('mainvote.db')
cursor = conn.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS poll (name)")

# 색상 테마 설정
COLORS = {
    'primary': '#1976D2',     # 주요 색상 (파란색)
    'secondary': '#2196F3',   # 보조 색상 (밝은 파란색)
    'accent': '#BBDEFB',      # 강조 색상 (매우 밝은 파란색)
    'background': '#F5F5F5',  # 배경 색상 (밝은 회색)
    'text': '#212121',        # 텍스트 색상 (진한 회색)
    'text_light': '#FFFFFF',  # 밝은 배경에 사용할 텍스트
    'success': '#4CAF50',     # 성공 색상 (녹색)
    'warning': '#FFC107',     # 경고 색상 (노란색)
    'error': '#F44336'        # 오류 색상 (빨간색)
}

# 글꼴 설정
FONTS = {
    'title': ('맑은 고딕', 16, 'bold'),
    'heading': ('맑은 고딕', 14, 'bold'),
    'normal': ('맑은 고딕', 11),
    'small': ('맑은 고딕', 9),
    'button': ('맑은 고딕', 11, 'bold')
}


# 스타일 설정
def setup_styles():
    style = ttk.Style()
    style.theme_use('clam')  # 기본 테마 설정

    # 버튼 스타일
    style.configure('TButton',
                    font=FONTS['button'],
                    background=COLORS['primary'],
                    foreground=COLORS['text_light'],
                    borderwidth=0,
                    focuscolor=COLORS['secondary'],
                    padding=(10, 5))

    # 라벨 스타일
    style.configure('TLabel',
                    font=FONTS['normal'],
                    background=COLORS['background'],
                    foreground=COLORS['text'])

    # 제목 라벨 스타일
    style.configure('Title.TLabel',
                    font=FONTS['title'],
                    background=COLORS['background'],
                    foreground=COLORS['primary'])

    # 콤보박스 스타일
    style.configure('TCombobox',
                    fieldbackground=COLORS['background'],
                    background=COLORS['primary'],
                    foreground=COLORS['text'])

    # 후보자 라디오 버튼
    style.configure('TRadiobutton',
                    font=('맑은 고딕', 12),
                    background=COLORS['background'])


# 사용자 정의 버튼 클래스
class CustomButton(tk.Button):
    def __init__(self, master=None, **kwargs):
        if 'bg' not in kwargs:
            kwargs['bg'] = COLORS['primary']
        if 'fg' not in kwargs:
            kwargs['fg'] = COLORS['text_light']
        if 'font' not in kwargs:
            kwargs['font'] = FONTS['button']
        if 'relief' not in kwargs:
            kwargs['relief'] = tk.RAISED
        if 'borderwidth' not in kwargs:
            kwargs['borderwidth'] = 1
        if 'padx' not in kwargs:
            kwargs['padx'] = 15
        if 'pady' not in kwargs:
            kwargs['pady'] = 5

        super().__init__(master, **kwargs)

        # 마우스 오버 효과
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self.config(bg=COLORS['secondary'])

    def on_leave(self, e):
        self.config(bg=COLORS['primary'])


# 메인 투표 프로그램 클래스
class VotingProgram:
    def __init__(self, root):
        self.root = root
        self.root.title("투표 프로그램")
        self.root.geometry("600x500")
        self.root.config(bg=COLORS['background'])

        # 스타일 설정
        setup_styles()

        # 프레임 생성
        self.setup_frames()

        # 메인 위젯 생성
        self.setup_widgets()

    def setup_frames(self):
        # 헤더 프레임
        self.header_frame = tk.Frame(self.root, bg=COLORS['primary'], height=70)
        self.header_frame.pack(fill=tk.X)

        # 본문 프레임
        self.content_frame = tk.Frame(self.root, bg=COLORS['background'])
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 푸터 프레임
        self.footer_frame = tk.Frame(self.root, bg=COLORS['background'], height=30)
        self.footer_frame.pack(fill=tk.X, side=tk.BOTTOM)

    def setup_widgets(self):
        # 헤더 내용
        tk.Label(self.header_frame,
                 text="투표 프로그램",
                 font=FONTS['title'],
                 bg=COLORS['primary'],
                 fg=COLORS['text_light']).pack(pady=15)

        # 본문 내용 - 그리드 레이아웃 사용
        tk.Label(self.content_frame,
                 text="안녕하세요! 투표 프로그램에 오신 것을 환영합니다",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['text']).grid(row=0, column=0, columnspan=3, pady=15)

        # 버튼들을 위한 프레임
        button_frame = tk.Frame(self.content_frame, bg=COLORS['background'])
        button_frame.grid(row=1, column=0, columnspan=3, pady=20)

        # 버튼 배치
        CustomButton(button_frame, text="새 투표 만들기", command=self.create_poll).pack(pady=10, fill=tk.X)
        CustomButton(button_frame, text="투표 참여하기", command=self.participate_poll).pack(pady=10, fill=tk.X)
        CustomButton(button_frame, text="투표 결과 보기", command=self.view_results).pack(pady=10, fill=tk.X)

        # 푸터 내용
        tk.Label(self.footer_frame,
                 text="개발자: samths@naver.com",
                 font=FONTS['small'],
                 bg=COLORS['background'],
                 fg=COLORS['text']).pack(side=tk.RIGHT, padx=10, pady=5)

        # 정보 버튼
        info_button = CustomButton(self.footer_frame, text="정보", command=self.show_about, bg=COLORS['accent'], fg=COLORS['text'])
        info_button.pack(side=tk.LEFT, padx=10, pady=5)

    def create_poll(self):
        create_window = tk.Toplevel(self.root)
        create_window.title("새 투표 만들기")
        create_window.geometry("520x400")
        create_window.config(bg=COLORS['background'])

        # 제목
        tk.Label(create_window,
                 text="새 투표 만들기",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['primary']).pack(pady=15)

        # 입력 프레임
        input_frame = tk.Frame(create_window, bg=COLORS['background'])
        input_frame.pack(padx=20, pady=10, fill=tk.X)

        # 투표 이름 입력
        tk.Label(input_frame, text="투표 이름:", font=FONTS['normal'], bg=COLORS['background']).grid(row=0, column=0, sticky=tk.W, pady=5)
        poll_name = tk.Entry(input_frame, width=30, font=FONTS['normal'])
        poll_name.grid(row=0, column=1, pady=5, padx=10)
        tk.Label(input_frame, text="(예: 학급 회장 선거)", font=FONTS['small'], bg=COLORS['background'], fg=COLORS['text']).grid(row=0, column=2, sticky=tk.W)

        # 후보자 입력
        tk.Label(input_frame, text="후보자 이름:", font=FONTS['normal'], bg=COLORS['background']).grid(row=1, column=0, sticky=tk.W, pady=5)
        candidates = tk.Text(input_frame, width=30, height=5, font=FONTS['normal'])
        candidates.grid(row=1, column=1, pady=5, padx=10)

        # 안내 메시지
        instruction_frame = tk.Frame(create_window, bg=COLORS['background'])
        instruction_frame.pack(padx=20, pady=5, fill=tk.X)

        tk.Label(instruction_frame,
                 text="참고: 각 후보자 이름을 쉼표(,)로 구분하여 입력하세요",
                 font=FONTS['small'],
                 bg=COLORS['background']).pack(anchor=tk.W)
        tk.Label(instruction_frame,
                 text="예: 홍길동, 김철수, 이영희",
                 font=FONTS['small'],
                 bg=COLORS['background']).pack(anchor=tk.W)

        # 버튼 프레임
        button_frame = tk.Frame(create_window, bg=COLORS['background'])
        button_frame.pack(pady=20)

        def proceed():
            pname = poll_name.get()  # 투표 이름
            can = candidates.get("1.0", tk.END).strip()  # 후보자 이름

            if pname == '':
                return messagebox.showerror('오류', '투표 이름을 입력하세요')
            elif can == '':
                return messagebox.showerror('오류', '후보자를 입력하세요')
            else:
                candidates_list = [c.strip() for c in can.split(',')]  # 후보자 목록

                # 데이터베이스에 저장
                command = 'INSERT INTO poll (name) VALUES (?);'
                cursor.execute(command, (pname,))
                conn.commit()

                pd = sqltor.connect(pname + '.db')  # 투표 데이터베이스
                pcursor = pd.cursor()  # 투표 커서
                pcursor.execute("""CREATE TABLE IF NOT EXISTS polling
                        (name TEXT, votes INTEGER)""")

                for candidate in candidates_list:
                    if candidate:  # 빈 문자열이 아닌 경우에만 추가
                        command = 'INSERT INTO polling (name, votes) VALUES (?, ?)'
                        data = (candidate, 0)
                        pcursor.execute(command, data)

                pd.commit()
                pd.close()
                messagebox.showinfo('성공!', '투표가 생성되었습니다')
                create_window.destroy()

        CustomButton(button_frame, text="생성하기", command=proceed).pack(side=tk.LEFT, padx=5)
        CustomButton(button_frame, text="취소", command=create_window.destroy, bg=COLORS['error']).pack(side=tk.LEFT, padx=5)

    def participate_poll(self):
        # 사용 가능한 투표 목록 가져오기
        cursor.execute('SELECT name FROM poll')
        data = cursor.fetchall()
        poll_names = ['-선택-']

        for i in range(len(data)):
            data1 = data[i]
            poll_names.append(data1[0])

        # 투표 선택 창
        select_window = tk.Toplevel(self.root)
        select_window.title("투표 참여하기")
        select_window.geometry("400x250")
        select_window.config(bg=COLORS['background'])

        # 제목
        tk.Label(select_window,
                 text="투표 선택",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['primary']).pack(pady=15)

        # 선택 프레임
        select_frame = tk.Frame(select_window, bg=COLORS['background'])
        select_frame.pack(padx=20, pady=10, fill=tk.X)

        tk.Label(select_frame, text="참여할 투표 선택:", font=FONTS['normal'], bg=COLORS['background']).grid(row=0, column=0, sticky=tk.W, pady=5)

        # 콤보박스
        poll_var = tk.StringVar()
        poll_select = ttk.Combobox(select_frame, values=poll_names, state='readonly', textvariable=poll_var, font=FONTS['normal'], width=20)
        poll_select.grid(row=0, column=1, pady=5, padx=10)
        poll_select.current(0)

        # 버튼 프레임
        button_frame = tk.Frame(select_window, bg=COLORS['background'])
        button_frame.pack(pady=20)

        def proceed():
            global plname
            plname = poll_var.get()

            if plname == '-선택-':
                return messagebox.showerror('오류', '투표를 선택하세요')
            else:
                select_window.destroy()
                self.open_poll_page(plname)

        CustomButton(button_frame, text="참여하기", command=proceed).pack(side=tk.LEFT, padx=5)
        CustomButton(button_frame, text="취소", command=select_window.destroy, bg=COLORS['error']).pack(side=tk.LEFT, padx=5)

    def open_poll_page(self, poll_name):
        # 투표 페이지 열기
        poll_window = tk.Toplevel(self.root)
        poll_window.title(f"{poll_name} - 투표")
        poll_window.geometry("450x350")
        poll_window.config(bg=COLORS['background'])

        # 제목
        tk.Label(poll_window,
                 text=f"{poll_name} - 투표",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['primary']).pack(pady=15)

        # 안내 라벨
        tk.Label(poll_window,
                 text="한 명의 후보를 선택하여 투표하세요",
                 font=FONTS['normal'],
                 bg=COLORS['background']).pack(pady=5)

        # 후보자 데이터 가져오기
        pd = sqltor.connect(poll_name + '.db')  # 투표 데이터베이스
        pcursor = pd.cursor()  # 투표 커서
        pcursor.execute('SELECT name FROM polling')
        data = pcursor.fetchall()

        names = []
        for i in range(len(data)):
            data1 = data[i]
            names.append(data1[0])

        # 후보자 프레임
        candidates_frame = tk.Frame(poll_window, bg=COLORS['background'])
        candidates_frame.pack(padx=20, pady=10, fill=tk.X)

        # 라디오 버튼
        choose = tk.StringVar()
        for i, name in enumerate(names):
            ttk.Radiobutton(candidates_frame,
                            text=name,
                            value=name,
                            variable=choose,
                            style='TRadiobutton').pack(anchor=tk.W, pady=3)

        # 버튼 프레임
        button_frame = tk.Frame(poll_window, bg=COLORS['background'])
        button_frame.pack(pady=20)

        def cast_vote():
            candidate = choose.get()

            if not candidate:
                return messagebox.showerror('오류', '후보자를 선택하세요')

            command = 'UPDATE polling SET votes = votes + 1 WHERE name = ?'
            pcursor.execute(command, (candidate,))
            pd.commit()
            messagebox.showinfo('성공!', '투표가 완료되었습니다')
            poll_window.destroy()

        CustomButton(button_frame, text="투표하기", command=cast_vote).pack(side=tk.LEFT, padx=5)
        CustomButton(button_frame, text="취소", command=poll_window.destroy, bg=COLORS['error']).pack(side=tk.LEFT, padx=5)

    def view_results(self):
        # 사용 가능한 투표 목록 가져오기
        cursor.execute('SELECT name FROM poll')
        data = cursor.fetchall()
        poll_names = ['-선택-']

        for i in range(len(data)):
            data1 = data[i]
            poll_names.append(data1[0])

        # 투표 선택 창
        result_window = tk.Toplevel(self.root)
        result_window.title("투표 결과")
        result_window.geometry("400x250")
        result_window.config(bg=COLORS['background'])

        # 제목
        tk.Label(result_window,
                 text="투표 결과 보기",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['primary']).pack(pady=15)

        # 선택 프레임
        select_frame = tk.Frame(result_window, bg=COLORS['background'])
        select_frame.pack(padx=20, pady=10, fill=tk.X)

        tk.Label(select_frame, text="결과를 볼 투표 선택:", font=FONTS['normal'], bg=COLORS['background']).grid(row=0, column=0, sticky=tk.W, pady=5)

        # 콤보박스
        poll_var = tk.StringVar()
        poll_select = ttk.Combobox(select_frame, values=poll_names, state='readonly', textvariable=poll_var, font=FONTS['normal'], width=20)
        poll_select.grid(row=0, column=1, pady=5, padx=10)
        poll_select.current(0)

        # 버튼 프레임
        button_frame = tk.Frame(result_window, bg=COLORS['background'])
        button_frame.pack(pady=20)

        def show_results():
            selected_poll = poll_var.get()

            if selected_poll == '-선택-':
                return messagebox.showerror('오류', '투표를 선택하세요')

            result_window.destroy()
            self.open_results_page(selected_poll)

        CustomButton(button_frame, text="결과 보기", command=show_results).pack(side=tk.LEFT, padx=5)
        CustomButton(button_frame, text="취소", command=result_window.destroy, bg=COLORS['error']).pack(side=tk.LEFT, padx=5)

    def open_results_page(self, poll_name):
        # 결과 페이지 열기
        results_window = tk.Toplevel(self.root)
        results_window.title(f"{poll_name} - 결과")
        results_window.geometry("500x400")
        results_window.config(bg=COLORS['background'])

        # 제목
        tk.Label(results_window,
                 text=f"{poll_name} - 투표 결과",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['primary']).pack(pady=15)

        # 결과 데이터 가져오기
        pd = sqltor.connect(poll_name + '.db')  # 투표 데이터베이스
        pcursor = pd.cursor()  # 투표 커서
        pcursor.execute('SELECT * FROM polling')
        data = pcursor.fetchall()

        # 결과 프레임
        results_frame = tk.Frame(results_window, bg=COLORS['background'])
        results_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        # 테이블 헤더
        tk.Label(results_frame, text="후보자", font=FONTS['normal'], bg=COLORS['primary'], fg=COLORS['text_light'], width=20, padx=10, pady=5).grid(row=0, column=0)
        tk.Label(results_frame, text="득표수", font=FONTS['normal'], bg=COLORS['primary'], fg=COLORS['text_light'], width=10, padx=10, pady=5).grid(row=0, column=1)

        # 결과 표시
        for i, row in enumerate(data):
            name, votes = row
            bg_color = COLORS['accent'] if i % 2 == 0 else COLORS['background']

            tk.Label(results_frame, text=name, font=FONTS['normal'], bg=bg_color, width=20, padx=10, pady=5).grid(row=i + 1, column=0, sticky="ew")
            tk.Label(results_frame, text=str(votes), font=FONTS['normal'], bg=bg_color, width=10, padx=10, pady=5).grid(row=i + 1, column=1, sticky="ew")

        # 버튼 프레임
        button_frame = tk.Frame(results_window, bg=COLORS['background'])
        button_frame.pack(pady=20)

        def show_graph():
            # 그래프 데이터 준비
            names = []
            votes = []

            for row in data:
                names.append(row[0])
                votes.append(row[1])

            # 그래프 생성
            plt.figure(figsize=(8, 6))
            plt.title(f'{poll_name} - 투표 결과', fontsize=16)

            # 파이 차트
            plt.pie(votes, labels=names, autopct='%1.1f%%', shadow=True, startangle=140, colors=plt.cm.Paired.colors)
            plt.axis('equal')  # 원형 파이 차트 비율 유지

            plt.show()

        CustomButton(button_frame, text="그래프로 보기", command=show_graph).pack(side=tk.LEFT, padx=5)
        CustomButton(button_frame, text="닫기", command=results_window.destroy, bg=COLORS['error']).pack(side=tk.LEFT, padx=5)

    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("프로그램 정보")
        about_window.geometry("500x360")
        about_window.config(bg=COLORS['background'])

        # 제목
        tk.Label(about_window,
                 text="투표 프로그램 정보",
                 font=FONTS['heading'],
                 bg=COLORS['background'],
                 fg=COLORS['primary']).pack(pady=15)

        # 정보 프레임
        info_frame = tk.Frame(about_window, bg=COLORS['background'])
        info_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        # 정보 내용
        tk.Label(info_frame,
                 text="투표 프로그램 Ver 2.0",
                 font=FONTS['normal'],
                 bg=COLORS['background']).pack(anchor=tk.W, pady=5)

        tk.Label(info_frame,
                 text="이 프로그램은 다양한 투표를 생성하고 관리할 수 있는 도구입니다.",
                 font=FONTS['normal'],
                 bg=COLORS['background']).pack(anchor=tk.W, pady=5)

        tk.Label(info_frame,
                 text="개발자: samths@naver.com",
                 font=FONTS['normal'],
                 bg=COLORS['background']).pack(anchor=tk.W, pady=5)

        tk.Label(info_frame,
                 text="GitHub: https://github.com/andrew-geeks",
                 font=FONTS['normal'],
                 bg=COLORS['background']).pack(anchor=tk.W, pady=2)

        tk.Label(info_frame,
                 text="Instagram: https://www.instagram.com/_andrewgeeks/",
                 font=FONTS['normal'],
                 bg=COLORS['background']).pack(anchor=tk.W, pady=2)

        # 버튼 프레임
        button_frame = tk.Frame(about_window, bg=COLORS['background'])
        button_frame.pack(pady=20)

        CustomButton(button_frame, text="확인", command=about_window.destroy).pack()


# 메인 실행 부분
if __name__ == "__main__":
    root = tk.Tk()
    app = VotingProgram(root)
    root.mainloop()