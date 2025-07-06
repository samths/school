"""
word_quiz.py     단어 관리 및 퀴즈 프로그램   Ver 1.0_250706
tkinter + Excel (words.xlsx) 사용
영어, 뜻 컬럼으로 구성
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import random
import os


class WordManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("단어 관리 및 퀴즈 프로그램")
        self.root.geometry("500x600")

        # Define the default font
        self.default_font = ('맑은 고딕', 11)

        # 데이터 저장용 변수
        self.df = pd.DataFrame(columns=['영어', '뜻'])
        self.file_path = "words.xlsx"

        # 퀴즈 관련 변수
        self.quiz_mode = False
        self.current_word = ""
        self.correct_count = 0
        self.total_count = 0
        self.quiz_words = []

        self.create_widgets()
        self.load_data()

        self.style = ttk.Style()
        self.style.configure('TButton', font=('맑은 고딕', 11))
        self.style.configure('Treeview.Heading', font=('맑은 고딕', 11))
        self.style.configure('Treeview', font=('맑은 고딕', 11))

    def create_widgets(self):
        # 노트북 (탭) 생성
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 퀴즈 탭
        self.quiz_frame = ttk.Frame(notebook)
        notebook.add(self.quiz_frame, text="퀴즈")

        # 단어 관리 탭
        self.manage_frame = ttk.Frame(notebook)
        notebook.add(self.manage_frame, text="단어 관리")

        self.create_manage_tab()
        self.create_quiz_tab()

    def create_manage_tab(self):
        # 상단 입력 프레임
        input_frame = ttk.LabelFrame(self.manage_frame, text="단어 입력/수정", padding="10")
        input_frame.pack(fill='x', padx=10, pady=5)

        # 영어 입력
        ttk.Label(input_frame, text="영어:", font=self.default_font).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.english_entry = ttk.Entry(input_frame, width=30, font=self.default_font)
        self.english_entry.grid(row=0, column=1, padx=5, pady=5)

        # 뜻 입력
        ttk.Label(input_frame, text="뜻:", font=self.default_font).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.meaning_entry = ttk.Entry(input_frame, width=30, font=self.default_font)
        self.meaning_entry.grid(row=1, column=1, padx=5, pady=5)

        # 버튼 프레임
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="추가", command=self.add_word).pack(side='left', padx=5)
        ttk.Button(button_frame, text="수정", command=self.update_word).pack(side='left', padx=5)
        ttk.Button(button_frame, text="삭제", command=self.delete_word).pack(side='left', padx=5)
        ttk.Button(button_frame, text="지우기", command=self.clear_entries).pack(side='left', padx=5)

        # 파일 관리 프레임
        file_frame = ttk.LabelFrame(self.manage_frame, text="파일 관리", padding="10")
        file_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(file_frame, text="파일 열기", command=self.open_file).pack(side='left', padx=5)
        ttk.Button(file_frame, text="저장", command=self.save_data).pack(side='left', padx=5)
        ttk.Button(file_frame, text="다른 이름으로 저장", command=self.save_as).pack(side='left', padx=5)

        # 단어 목록 프레임
        list_frame = ttk.LabelFrame(self.manage_frame, text="단어 목록", padding="10")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # 트리뷰 생성 - 컬럼 설정 수정
        self.tree = ttk.Treeview(list_frame, columns=('영어', '뜻'), show='headings', height=15)

        # 컬럼 설정
        self.tree.heading('영어', text='영어')
        self.tree.heading('뜻', text='뜻')
        self.tree.column('영어', width=200, anchor='w')
        self.tree.column('뜻', width=350, anchor='w')

        # 스크롤바 - 수직 및 수평 스크롤바 추가
        v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        # 트리뷰와 스크롤바 배치
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')

        # 그리드 가중치 설정
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # 트리뷰 이벤트 바인딩
        self.tree.bind('<Double-1>', self.on_item_select)
        self.tree.bind('<Button-1>', self.on_item_click)

    def create_quiz_tab(self):
        # 퀴즈 설정 프레임
        quiz_setup_frame = ttk.LabelFrame(self.quiz_frame, text="퀴즈 설정", padding="10")
        quiz_setup_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(quiz_setup_frame, text="퀴즈 시작", command=self.start_quiz).pack(side='left', padx=5)
        ttk.Button(quiz_setup_frame, text="퀴즈 중단", command=self.stop_quiz).pack(side='left', padx=5)

        # 퀴즈 진행 프레임
        quiz_main_frame = ttk.LabelFrame(self.quiz_frame, text="퀴즈", padding="20")
        quiz_main_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # 단어 표시
        self.word_label = ttk.Label(quiz_main_frame, text="퀴즈를 시작해주세요.", font=('맑은 고딕', 16, 'bold'))
        self.word_label.pack(pady=20)

        # 답 입력
        self.answer_entry = ttk.Entry(quiz_main_frame, width=40, font=('맑은 고딕', 12))
        self.answer_entry.pack(pady=10)
        self.answer_entry.bind('<Return>', lambda event: self.check_answer())

        # 제출 버튼
        ttk.Button(quiz_main_frame, text="제출", command=self.check_answer).pack(pady=10)

        # 점수 표시
        self.score_label = ttk.Label(quiz_main_frame, text="점수: 0/0", font=('맑은 고딕', 12))
        self.score_label.pack(pady=10)

        # 결과 표시
        self.result_label = ttk.Label(quiz_main_frame, text="", font=('맑은 고딕', 11))
        self.result_label.pack(pady=5)

    def load_data(self):
        """데이터 로드"""
        try:
            if os.path.exists(self.file_path):
                self.df = pd.read_excel(self.file_path)
                # 컬럼명 확인 및 수정
                if '영어' not in self.df.columns or '뜻' not in self.df.columns:
                    messagebox.showwarning("파일 형식 오류", "Excel 파일에 '영어', '뜻' 컬럼이 없습니다.\n빈 데이터프레임을 생성합니다.")
                    self.df = pd.DataFrame(columns=['영어', '뜻'])
                # NaN 값 처리
                self.df = self.df.fillna('')
            else:
                self.df = pd.DataFrame(columns=['영어', '뜻'])
            self.update_tree()
            print(f"데이터 로드 완료: {len(self.df)}개 단어")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는 중 오류가 발생했습니다: {str(e)}")
            self.df = pd.DataFrame(columns=['영어', '뜻'])

    def save_data(self):
        """데이터 저장"""
        try:
            self.df.to_excel(self.file_path, index=False)
            messagebox.showinfo("저장 완료", f"파일이 저장되었습니다.\n({self.file_path})")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 저장하는 중 오류가 발생했습니다: {str(e)}")

    def open_file(self):
        """파일 열기"""
        file_path = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.file_path = file_path
                self.load_data()
                messagebox.showinfo("파일 열기", f"파일이 성공적으로 열렸습니다.\n{len(self.df)}개 단어가 로드되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"파일을 여는 중 오류가 발생했습니다: {str(e)}")

    def save_as(self):
        """다른 이름으로 저장"""
        file_path = filedialog.asksaveasfilename(
            title="다른 이름으로 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.file_path = file_path
                self.save_data()
            except Exception as e:
                messagebox.showerror("오류", f"파일을 저장하는 중 오류가 발생했습니다: {str(e)}")

    def update_tree(self):
        """트리뷰 업데이트"""
        # 기존 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 새 데이터 추가
        for index, row in self.df.iterrows():
            english = str(row['영어']) if pd.notna(row['영어']) else ''
            meaning = str(row['뜻']) if pd.notna(row['뜻']) else ''
            self.tree.insert('', 'end', values=(english, meaning))

        print(f"트리뷰 업데이트 완료: {len(self.df)}개 항목")

    def add_word(self):
        """단어 추가"""
        english = self.english_entry.get().strip()
        meaning = self.meaning_entry.get().strip()

        if not english or not meaning:
            messagebox.showwarning("입력 오류", "영어와 뜻을 모두 입력해주세요.")
            return

        # 중복 확인
        if not self.df.empty and english in self.df['영어'].values:
            messagebox.showwarning("중복 오류", "이미 존재하는 단어입니다.")
            return

        # 데이터 추가
        new_row = pd.DataFrame({'영어': [english], '뜻': [meaning]})
        self.df = pd.concat([self.df, new_row], ignore_index=True)

        self.update_tree()
        self.clear_entries()
        messagebox.showinfo("추가 완료", "단어가 추가되었습니다.")

    def update_word(self):
        """실제 단어 수정 처리"""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "수정할 단어를 먼저 선택해주세요.")
            return

        english = self.english_entry.get().strip()
        meaning = self.meaning_entry.get().strip()

        if not english or not meaning:
            messagebox.showwarning("입력 오류", "영어와 뜻을 모두 입력해주세요.")
            return

        # 선택된 항목의 원래 내용
        item_values = self.tree.item(selected_item[0])['values']
        old_english = item_values[0]

        # 최종 수정 확인
        if messagebox.askyesno("수정 확인",
                               f"다음과 같이 수정하시겠습니까?\n\n이전: {old_english} → {item_values[1]}\n수정: {english} → {meaning}"):
            # 데이터 수정
            try:
                mask = self.df['영어'] == old_english
                if mask.any():
                    index = self.df[mask].index[0]
                    self.df.loc[index, '영어'] = english
                    self.df.loc[index, '뜻'] = meaning

                    self.update_tree()
                    self.clear_entries()
                    messagebox.showinfo("수정 완료", f"단어가 수정되었습니다.\n{old_english} → {english}")
                else:
                    messagebox.showerror("오류", "수정할 단어를 찾을 수 없습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"수정 중 오류가 발생했습니다: {str(e)}")

    def delete_word(self):
        """단어 삭제"""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "삭제할 단어를 먼저 선택해주세요.")
            return

        try:
            # 선택된 항목의 내용을 가져와서 입력 필드에 표시
            item_values = self.tree.item(selected_item[0])['values']
            if len(item_values) >= 2:
                english = item_values[0]
                meaning = item_values[1]

                # 입력 필드에 내용 표시
                self.english_entry.delete(0, tk.END)
                self.meaning_entry.delete(0, tk.END)
                self.english_entry.insert(0, english)
                self.meaning_entry.insert(0, meaning)

                # 삭제 확인 대화상자에 내용 표시
                if messagebox.askyesno("삭제 확인",
                                       f"다음 단어를 삭제하시겠습니까?\n\n영어: {english}\n뜻: {meaning}"):
                    # 데이터 삭제
                    self.df = self.df[self.df['영어'] != english].reset_index(drop=True)

                    self.update_tree()
                    self.clear_entries()
                    messagebox.showinfo("삭제 완료", f"단어가 삭제되었습니다.\n({english})")
                else:
                    # 삭제 취소시 입력 필드 유지
                    pass
            else:
                messagebox.showerror("오류", "선택한 항목의 데이터를 읽을 수 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"삭제 중 오류가 발생했습니다: {str(e)}")

    def clear_entries(self):
        """입력 필드 지우기"""
        self.english_entry.delete(0, tk.END)
        self.meaning_entry.delete(0, tk.END)

    def on_item_click(self, event):
        """트리뷰 항목 클릭 시 (단일 클릭)"""
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item[0])['values']
            if len(item_values) >= 2:
                self.english_entry.delete(0, tk.END)
                self.meaning_entry.delete(0, tk.END)
                self.english_entry.insert(0, item_values[0])
                self.meaning_entry.insert(0, item_values[1])

    def on_item_select(self, event):
        """트리뷰 항목 더블클릭 시"""
        self.on_item_click(event)

    def start_quiz(self):
        """퀴즈 시작"""
        if len(self.df) == 0:
            messagebox.showwarning("퀴즈 오류", "단어가 없습니다. 먼저 단어를 추가해주세요.")
            return

        self.quiz_mode = True
        self.quiz_words = self.df.sample(frac=1).reset_index(drop=True)  # 랜덤 셔플
        self.current_word = ""
        self.correct_count = 0
        self.total_count = 0

        self.update_score()
        self.next_word()

    def stop_quiz(self):
        """퀴즈 중단"""
        self.quiz_mode = False
        self.word_label.config(text="퀴즈를 시작해주세요.")
        self.answer_entry.delete(0, tk.END)
        self.result_label.config(text="")
        self.score_label.config(text="점수: 0/0")

    def next_word(self):
        """다음 단어"""
        if len(self.quiz_words) > 0:
            current_row = self.quiz_words.iloc[0]
            self.current_word = current_row['영어']
            self.word_label.config(text=f"'{self.current_word}'의 뜻은?")
            self.quiz_words = self.quiz_words.iloc[1:]  # 첫 번째 행 제거
            self.answer_entry.delete(0, tk.END)
            self.answer_entry.focus()
        else:
            self.quiz_complete()

    def check_answer(self):
        """정답 확인"""
        if not self.quiz_mode or not self.current_word:
            return

        user_answer = self.answer_entry.get().strip()
        correct_answer = self.df[self.df['영어'] == self.current_word]['뜻'].iloc[0]

        self.total_count += 1

        if user_answer.lower() == correct_answer.lower():
            self.correct_count += 1
            self.result_label.config(text="정답입니다! ✓", foreground="green")
        else:
            self.result_label.config(text=f"틀렸습니다. 정답: {correct_answer}", foreground="red")

        self.update_score()
        self.root.after(1500, self.next_word)  # 1.5초 후 다음 단어

    def update_score(self):
        """점수 업데이트"""
        self.score_label.config(text=f"점수: {self.correct_count}/{self.total_count}")

    def quiz_complete(self):
        """퀴즈 완료"""
        self.quiz_mode = False
        accuracy = (self.correct_count / self.total_count * 100) if self.total_count > 0 else 0

        result_message = f"퀴즈 완료!\n총 {self.total_count}문제 중 {self.correct_count}문제 정답\n정답률: {accuracy:.1f}%"
        messagebox.showinfo("퀴즈 완료", result_message)

        self.word_label.config(text="퀴즈가 완료되었습니다.")
        self.result_label.config(text="")


if __name__ == "__main__":
    root = tk.Tk()
    app = WordManagerApp(root)
    root.mainloop()