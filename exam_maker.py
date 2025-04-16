"""
exam_maker.py  시험 문제 은행  Ver 1.1_250412
"""
import tkinter as tk
from tkinter import messagebox, font, ttk, simpledialog
import random
import csv
import os
import pandas as pd


class QuizCreator:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Exam Creator")
        self.window.geometry("800x600")

        default_font = font.Font(family="맑은 고딕", size=11)
        self.window.option_add("*Font", default_font)

        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill="both", expand=True)

        self.username = ""  # 사용자 이름 저장 변수

        # 결과 파일이 없으면 생성
        if not os.path.exists("exam_results.csv"):
            with open("exam_results.csv", "w", newline='', encoding="utf-8") as f:
                pass

        self.create_entry_tab()
        self.create_manage_tab()
        self.create_statistics_tab()

    def create_entry_tab(self):
        self.entry_tab = tk.Frame(self.notebook)
        self.notebook.add(self.entry_tab, text="문제 작성")

        tk.Label(self.entry_tab, text="문제 :").pack()
        self.question_entry = tk.Text(self.entry_tab, height=5, width=90)
        self.question_entry.pack()

        self.options_entries = []
        for i in range(1, 6):
            label = tk.Label(self.entry_tab, pady=4, text=f"보기 {i}:")
            label.pack()
            entry = tk.Entry(self.entry_tab, width=90)
            entry.pack()
            self.options_entries.append(entry)

        answer_frame = tk.Frame(self.entry_tab)
        answer_frame.pack(pady=5)
        tk.Label(answer_frame, text="정답 (예: 1 또는 1,3):").grid(row=0, column=0)
        self.correct_answer_entry = tk.Entry(answer_frame, width=15)
        self.correct_answer_entry.grid(row=0, column=1, padx=5)

        tk.Label(answer_frame, text="난이도 (1~5):").grid(row=0, column=2)
        self.difficulty_entry = tk.Entry(answer_frame, width=5)
        self.difficulty_entry.grid(row=0, column=3)

        button_frame = tk.Frame(self.entry_tab)
        button_frame.pack(pady=10)
        tk.Button(button_frame, text="문제 저장", width=15, command=self.save_question).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="입력 지우기", width=15, command=self.clear_fields).grid(row=0, column=1, padx=5)

    def create_manage_tab(self):
        self.manage_tab = tk.Frame(self.notebook)
        self.notebook.add(self.manage_tab, text="문제 관리")

        # 상단 버튼 프레임 생성
        button_frame = tk.Frame(self.manage_tab)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="문제 목록 보기", command=self.show_question_list).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="문제 검색", command=self.search_question).grid(row=0, column=1, padx=5)
        tk.Button(button_frame, text="시험 시작", command=self.request_username_before_quiz).grid(row=0, column=2, padx=5)

        # 검색 프레임 생성
        self.search_frame = tk.Frame(self.manage_tab)
        self.search_frame.pack(fill="x", pady=5)
        tk.Label(self.search_frame, text="검색어:").pack(side="left", padx=5)
        self.search_entry = tk.Entry(self.search_frame, width=30)
        self.search_entry.pack(side="left", padx=5)
        tk.Button(self.search_frame, text="검색", command=self.perform_search).pack(side="left", padx=5)
        tk.Button(self.search_frame, text="전체 보기", command=self.show_question_list).pack(side="left", padx=5)

        # TreeView 영역 생성 (좌우 10px 여백 추가)
        tree_outer_frame = tk.Frame(self.manage_tab, padx=10)
        tree_outer_frame.pack(fill="both", expand=True, pady=5)

        tree_frame = tk.Frame(tree_outer_frame)
        tree_frame.pack(fill="both", expand=True)

        # TreeView 생성
        self.questions_tree = ttk.Treeview(tree_frame, columns=list(range(8)), show="headings")
        columns = ["문제", "1", "2", "3", "4", "5", "정답", "난이도"]
        for i, col in enumerate(columns):
            self.questions_tree.heading(i, text=col)
            if i == 0:  # 문제 컬럼 더 넓게
                self.questions_tree.column(i, width=200)
            else:
                self.questions_tree.column(i, width=70)

        # 스크롤바 추가
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.questions_tree.yview)
        self.questions_tree.configure(yscrollcommand=tree_scroll.set)

        # 배치
        self.questions_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        # 하단 버튼 프레임
        action_frame = tk.Frame(self.manage_tab)
        action_frame.pack(pady=5)
        tk.Button(action_frame, text="선택 삭제", command=self.delete_selected).pack(side="left", padx=5)
        tk.Button(action_frame, text="선택 편집", command=self.edit_selected).pack(side="left", padx=5)

        # 상태 표시줄
        self.status_var = tk.StringVar()
        self.status_var.set("준비")
        self.status_label = tk.Label(self.manage_tab, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def create_statistics_tab(self):
        self.stats_tab = tk.Frame(self.notebook)
        self.notebook.add(self.stats_tab, text="통계")

        tk.Label(self.stats_tab, text="사용자 이름 입력:").pack(pady=(20, 5))
        self.username_entry = tk.Entry(self.stats_tab, width=25)
        self.username_entry.pack()

        tk.Button(self.stats_tab, text="정답 통계 보기", command=self.show_statistics).pack(pady=10)

        self.stats_result_label = tk.Label(self.stats_tab, text="", font=("맑은 고딕", 12), fg="blue")
        self.stats_result_label.pack(pady=10)

        # 통계 세부정보 표시용 TreeView
        tk.Label(self.stats_tab, text="세부 결과:").pack(pady=(10, 5))
        stats_tree_frame = tk.Frame(self.stats_tab)
        stats_tree_frame.pack(fill="both", expand=True, padx=10)

        self.stats_tree = ttk.Treeview(stats_tree_frame, columns=["문제", "선택", "정답", "결과"], show="headings")
        self.stats_tree.heading("문제", text="문제")
        self.stats_tree.heading("선택", text="선택")
        self.stats_tree.heading("정답", text="정답")
        self.stats_tree.heading("결과", text="결과")

        self.stats_tree.column("문제", width=300)
        self.stats_tree.column("선택", width=100)
        self.stats_tree.column("정답", width=100)
        self.stats_tree.column("결과", width=80)

        stats_scroll = ttk.Scrollbar(stats_tree_frame, orient="vertical", command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=stats_scroll.set)

        self.stats_tree.pack(side="left", fill="both", expand=True)
        stats_scroll.pack(side="right", fill="y")

    def get_question_data(self):
        question = self.question_entry.get("1.0", tk.END).strip()
        answers = [entry.get().strip() for entry in self.options_entries]
        correct = self.correct_answer_entry.get().strip()
        difficulty = self.difficulty_entry.get().strip()
        return question, answers, correct, difficulty

    def save_question(self):
        question, answers, correct, difficulty = self.get_question_data()
        if not question or not all(answers) or not difficulty.isdigit() or not (1 <= int(difficulty) <= 5):
            messagebox.showerror("오류", "모든 항목을 작성하고 난이도는 1~5 사이 숫자여야 합니다.")
            return

        # 정답 검증
        correct_answers = correct.replace(" ", "").split(',')
        for c in correct_answers:
            if c not in ['1', '2', '3', '4', '5']:
                messagebox.showerror("오류", "정답은 1~5 범위에서 ,로 구분 입력하세요.")
                return

        # 중복 검사
        if os.path.exists("exam.csv"):
            with open("exam.csv", encoding="utf-8") as f:
                if question in [row[0] for row in csv.reader(f)]:
                    messagebox.showerror("중복", "이미 존재하는 문제입니다.")
                    return

        # CSV 파일에 저장
        try:
            with open("exam.csv", "a", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([question] + answers + [correct, difficulty])
            self.update_excel_file()

            self.clear_fields()
            messagebox.showinfo("성공", "문제가 저장되었습니다.")

            # 문제 저장 후 자동으로 목록 갱신
            self.show_question_list()
        except Exception as e:
            messagebox.showerror("저장 오류", f"문제 저장 중 오류가 발생했습니다: {e}")

    def update_excel_file(self):
        try:
            if os.path.exists("exam.csv"):
                df = pd.read_csv("exam.csv", header=None)
                df.columns = ['Question', '1', '2', '3', '4', '5', 'Answer', 'Difficulty']
                df.to_excel("exam.xlsx", index=False)
        except Exception as e:
            messagebox.showerror("Excel 변환 오류", f"Excel 파일 생성 중 오류가 발생했습니다: {e}")

    def clear_fields(self):
        self.question_entry.delete("1.0", tk.END)
        for entry in self.options_entries:
            entry.delete(0, tk.END)
        self.correct_answer_entry.delete(0, tk.END)
        self.difficulty_entry.delete(0, tk.END)

    def show_question_list(self):
        if not os.path.exists("exam.csv") or os.path.getsize("exam.csv") == 0:
            self.status_var.set("저장된 문제가 없습니다.")
            messagebox.showinfo("정보", "저장된 문제가 없습니다.")
            return

        # TreeView 내용 초기화
        self.questions_tree.delete(*self.questions_tree.get_children())

        try:
            # 문제 데이터 로드
            with open("exam.csv", encoding="utf-8") as f:
                data = list(csv.reader(f))

            # TreeView에 데이터 추가
            for row in data:
                if len(row) >= 8:  # 올바른 데이터 형식인지 확인
                    self.questions_tree.insert("", tk.END, values=row)

            # 관리 탭으로 전환
            self.notebook.select(self.manage_tab)
            self.status_var.set(f"전체 문제 {len(data)}개 표시됨")
        except Exception as e:
            messagebox.showerror("데이터 로드 오류", f"문제 목록을 불러오는 중 오류가 발생했습니다: {e}")

    def delete_selected(self):
        selected = self.questions_tree.selection()
        if not selected:
            messagebox.showinfo("알림", "삭제할 문제를 선택하세요.")
            return

        if not messagebox.askyesno("확인", "선택한 문제를 삭제하시겠습니까?"):
            return

        try:
            # 데이터 로드
            with open("exam.csv", encoding="utf-8") as f:
                data = list(csv.reader(f))

            # 선택한 항목 삭제
            index = self.questions_tree.index(selected[0])
            del data[index]

            # CSV 파일 업데이트
            with open("exam.csv", "w", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(data)

            self.update_excel_file()

            # TreeView 갱신
            self.questions_tree.delete(*self.questions_tree.get_children())
            for row in data:
                self.questions_tree.insert("", tk.END, values=row)

            messagebox.showinfo("삭제됨", "문제가 삭제되었습니다.")
            self.status_var.set(f"문제 삭제됨. 전체 {len(data)}개 문제 남음")
        except Exception as e:
            messagebox.showerror("삭제 오류", f"문제 삭제 중 오류가 발생했습니다: {e}")

    def edit_selected(self):
        selected = self.questions_tree.selection()
        if not selected:
            messagebox.showinfo("알림", "편집할 문제를 선택하세요.")
            return

        try:
            # 선택한 항목의 데이터 가져오기
            values = self.questions_tree.item(selected[0])['values']
            if len(values) < 8:
                messagebox.showerror("데이터 오류", "선택한 문제의 데이터가 올바르지 않습니다.")
                return

            # 작성 탭으로 전환 및 데이터 입력
            self.notebook.select(self.entry_tab)
            self.question_entry.delete("1.0", tk.END)
            self.question_entry.insert(tk.END, values[0])

            for i, entry in enumerate(self.options_entries):
                entry.delete(0, tk.END)
                entry.insert(0, values[i + 1])

            self.correct_answer_entry.delete(0, tk.END)
            self.correct_answer_entry.insert(0, values[6])

            self.difficulty_entry.delete(0, tk.END)
            self.difficulty_entry.insert(0, values[7])

            # 원본 데이터에서 해당 문제 삭제
            with open("exam.csv", encoding="utf-8") as f:
                data = list(csv.reader(f))

            del data[self.questions_tree.index(selected[0])]

            with open("exam.csv", "w", newline='', encoding="utf-8") as f:
                csv.writer(f).writerows(data)

            self.update_excel_file()

            # TreeView 갱신
            self.show_question_list()
            self.status_var.set("문제 편집 모드")
        except Exception as e:
            messagebox.showerror("편집 오류", f"문제 편집 중 오류가 발생했습니다: {e}")

    def search_question(self):
        if not os.path.exists("exam.csv") or os.path.getsize("exam.csv") == 0:
            messagebox.showinfo("정보", "문제가 없습니다.")
            return

        # 관리 탭으로 전환
        self.notebook.select(self.manage_tab)

        # 검색창에 포커스
        self.search_entry.focus_set()
        self.status_var.set("검색어를 입력하세요")

    def perform_search(self):
        keyword = self.search_entry.get().strip()
        if not keyword:
            messagebox.showinfo("알림", "검색어를 입력하세요.")
            return

        if not os.path.exists("exam.csv") or os.path.getsize("exam.csv") == 0:
            messagebox.showinfo("정보", "문제가 없습니다.")
            return

        try:
            # 데이터 로드
            with open("exam.csv", encoding="utf-8") as f:
                data = list(csv.reader(f))

            # 검색 결과 필터링 (문제와 보기 옵션 모두 검색)
            results = []
            for row in data:
                if any(keyword in col for col in row[:6]):  # 문제와 보기 옵션만 검색
                    results.append(row)

            # TreeView 내용 초기화
            self.questions_tree.delete(*self.questions_tree.get_children())

            # 검색 결과 추가
            for row in results:
                self.questions_tree.insert("", tk.END, values=row)

            if results:
                self.status_var.set(f"'{keyword}' 검색 결과: {len(results)}개 문제 찾음")
            else:
                self.status_var.set(f"'{keyword}' 검색 결과가 없습니다.")
        except Exception as e:
            messagebox.showerror("검색 오류", f"검색 중 오류가 발생했습니다: {e}")

    def request_username_before_quiz(self):
        if not os.path.exists("exam.csv") or os.path.getsize("exam.csv") == 0:
            messagebox.showinfo("정보", "저장된 문제가 없습니다.")
            return

        self.username = simpledialog.askstring("사용자 이름 입력", "시험을 보기 전에 이름을 입력하세요:")
        if self.username:
            self.start_full_quiz()
        else:
            messagebox.showerror("오류", "이름이 입력되지 않았습니다.")

    def start_full_quiz(self):
        try:
            with open("exam.csv", encoding="utf-8") as f:
                self.quiz_questions = list(csv.reader(f))

            if not self.quiz_questions:
                messagebox.showinfo("정보", "문제가 없습니다.")
                return

            random.shuffle(self.quiz_questions)  # 문제 순서 섞기
            self.current_q_index = 0
            self.quiz_score = 0
            self.quiz_next()
        except Exception as e:
            messagebox.showerror("시험 시작 오류", f"시험을 시작하는 중 오류가 발생했습니다: {e}")

    def quiz_next(self):
        if self.current_q_index >= len(self.quiz_questions):
            messagebox.showinfo("완료", f"시험이 끝났습니다.\n정답: {self.quiz_score}/{len(self.quiz_questions)}")
            self.notebook.select(self.stats_tab)  # 통계 탭으로 자동 이동
            self.username_entry.delete(0, tk.END)
            self.username_entry.insert(0, self.username)
            self.show_statistics()  # 통계 자동 표시
            return

        try:
            q = self.quiz_questions[self.current_q_index]
            q_window = tk.Toplevel(self.window)
            q_window.title(f"문제 {self.current_q_index + 1}")
            q_window.geometry("600x400")
            q_window.transient(self.window)  # 부모 창에 종속
            q_window.grab_set()  # 모달 창으로 설정

            # 문제 프레임
            question_frame = tk.Frame(q_window)
            question_frame.pack(fill="both", expand=True, padx=20, pady=10)

            # 문제 텍스트
            tk.Label(question_frame, text=q[0], wraplength=550,
                     font=("맑은 고딕", 12, "bold"), justify="left").pack(anchor="w", pady=10)

            # 보기 프레임
            options_frame = tk.Frame(q_window)
            options_frame.pack(fill="both", expand=True, padx=20)

            var_list = []
            for i in range(1, 6):
                var = tk.IntVar()
                cb = tk.Checkbutton(options_frame, text=f"{i}) {q[i]}", variable=var,
                                    font=("맑은 고딕", 11), wraplength=500, justify="left")
                cb.pack(anchor="w", pady=3)
                var_list.append(var)

            # 버튼 프레임
            button_frame = tk.Frame(q_window)
            button_frame.pack(fill="x", padx=20, pady=10)

            # 남은 문제 표시
            tk.Label(button_frame,
                     text=f"진행: {self.current_q_index + 1}/{len(self.quiz_questions)}").pack(side="left")

            def submit():
                try:
                    selected = [str(i + 1) for i, v in enumerate(var_list) if v.get() == 1]
                    if not selected:
                        messagebox.showwarning("경고", "하나 이상의 답을 선택하세요.", parent=q_window)
                        return

                    correct = [c.strip() for c in q[6].split(',')]
                    is_correct = sorted(selected) == sorted(correct)
                    result = "정답" if is_correct else "오답"

                    if is_correct:
                        self.quiz_score += 1

                    # 결과 저장
                    with open("exam_results.csv", "a", newline='', encoding="utf-8") as f:
                        writer = csv.writer(f)
                        writer.writerow([self.username, q[0], ",".join(selected), ",".join(correct), result])

                    # 정답 표시 (1초 후 닫기)
                    result_text = "정답입니다!" if is_correct else f"오답입니다. 정답은 {', '.join(correct)}번 입니다."
                    result_color = "green" if is_correct else "red"

                    result_label = tk.Label(q_window, text=result_text, font=("맑은 고딕", 12, "bold"),
                                            fg=result_color)
                    result_label.pack(pady=10)

                    submit_btn.config(state="disabled")
                    q_window.after(1500, lambda: q_window.destroy())  # 1.5초 후 창 닫기

                    self.current_q_index += 1
                    self.window.after(1600, self.quiz_next)  # 1.6초 후 다음 문제 표시
                except Exception as e:
                    messagebox.showerror("제출 오류", f"답안 제출 중 오류가 발생했습니다: {e}", parent=q_window)

            submit_btn = tk.Button(button_frame, text="제출", command=submit, width=10)
            submit_btn.pack(side="right")

        except Exception as e:
            messagebox.showerror("문제 표시 오류", f"문제를 표시하는 중 오류가 발생했습니다: {e}")
            self.current_q_index += 1
            self.quiz_next()  # 다음 문제로 이동

    def show_statistics(self):
        name = self.username_entry.get().strip()
        if not name:
            self.stats_result_label.config(text="사용자 이름을 입력하세요.")
            return

        try:
            if not os.path.exists("exam_results.csv") or os.path.getsize("exam_results.csv") == 0:
                self.stats_result_label.config(text="결과 파일이 존재하지 않습니다.")
                return

            # 트리뷰 초기화
            self.stats_tree.delete(*self.stats_tree.get_children())

            df = pd.read_csv("exam_results.csv", header=None, names=["사용자", "문제", "선택", "정답", "결과"])
            user_df = df[df["사용자"] == name]

            if user_df.empty:
                self.stats_result_label.config(text=f"{name}님의 기록이 없습니다.")
                return

            total = len(user_df)
            correct = len(user_df[user_df["결과"] == "정답"])
            accuracy = correct / total * 100

            # 난이도별 통계
            difficulty_stats = ""
            if os.path.exists("exam.csv"):
                try:
                    exam_df = pd.read_csv("exam.csv", header=None)
                    exam_df.columns = ["문제", "보기1", "보기2", "보기3", "보기4", "보기5", "정답", "난이도"]

                    # 사용자가 풀었던 문제의 난이도 정보 추출
                    merged = pd.merge(user_df, exam_df, left_on="문제", right_on="문제", how="left")
                    if not merged.empty and "난이도" in merged.columns:
                        # 난이도별 문제 수 및 정답률 계산
                        difficulty_grouped = merged.groupby("난이도").agg(
                            문제수=("난이도1", "count"),
                            정답수=("결과", lambda x: (x == "정답").sum())
                        )
                        difficulty_grouped["정답률"] = (difficulty_grouped["정답수"] / difficulty_grouped["문제수"] * 100).round(1)

                        # 문자열로 변환
                        difficulty_stats = "\n난이도별 통계:\n"
                        for idx, row in difficulty_grouped.iterrows():
                            difficulty_stats += f"난이도 {idx}: {row['문제수']}문제 중 {row['정답수']}문제 정답 ({row['정답률']}%)\n"
                except:
                    pass

            result_text = (
                f"{name}님의 정답 통계 총 응시 문제 수: {total} 정답 수: {correct} 정답률: {accuracy:.1f}%"
                f"{difficulty_stats}"
            )
            self.stats_result_label.config(text=result_text)

            # 세부 결과를 트리뷰에 추가
            for idx, row in user_df.iterrows():
                # 문제 내용이 너무 길면 잘라서 표시
                question = row["문제"]
                if len(question) > 50:
                    question = question[:50] + "..."

                self.stats_tree.insert("", tk.END, values=(
                    question,
                    row["선택"],
                    row["정답"],
                    row["결과"]
                ))

        except Exception as e:
            self.stats_result_label.config(text=f"통계 데이터 로드 중 오류 발생: {e}")

    def run(self):
        # CSV 파일이 없을 경우 생성
        if not os.path.exists("exam.csv"):
            with open("exam.csv", "w", newline='', encoding="utf-8") as f:
                pass

        self.window.mainloop()


if __name__ == "__main__":
    QuizCreator().run()