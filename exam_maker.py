"""
exam_maker.py  시험 문제 출제  Ver 1.0_250412
"""
import tkinter as tk
from tkinter import messagebox, font, ttk
import random
import csv
import os
import pandas as pd


class QuizCreator:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Exam Creator")
        self.window.geometry("600x600")

        default_font = font.Font(family="맑은 고딕", size=11)
        self.window.option_add("*Font", default_font)

        # 문제 입력
        self.question_label = tk.Label(self.window, text="문제 :")
        self.question_label.pack()
        self.question_entry = tk.Text(self.window, height=5, width=60)
        self.question_entry.pack()

        # 보기들
        self.options_entries = []
        for i in range(1, 6):
            label = tk.Label(self.window, pady=4, text=f"보기 {i}:")
            label.pack()
            entry = tk.Entry(self.window, width=60)
            entry.pack()
            self.options_entries.append(entry)

        # 정답 입력 (1,3 식으로 복수 입력 가능)
        answer_frame = tk.Frame(self.window)
        answer_frame.pack(pady=10)
        self.correct_answer_label = tk.Label(answer_frame, text="정답 (예: 1 또는 1,3):")
        self.correct_answer_label.grid(row=0, column=0, padx=5)
        self.correct_answer_entry = tk.Entry(answer_frame, width=15)
        self.correct_answer_entry.grid(row=0, column=1)

        # 버튼들
        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="문제 저장", width=12, command=self.save_question).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="문제 목록", width=12, command=self.show_question_list).grid(row=0, column=1, padx=5)
        tk.Button(button_frame, text="문제 검색", width=12, command=self.search_question).grid(row=0, column=2, padx=5)
        tk.Button(button_frame, text="랜덤 시험 시작", width=12, command=self.start_quiz).grid(row=0, column=3, padx=5)
        tk.Button(self.window, text="종료", width=12, command=self.window.destroy).pack(pady=5)

    def get_question_data(self):
        question = self.question_entry.get("1.0", tk.END).strip()
        answers = [entry.get().strip() for entry in self.options_entries]
        correct = self.correct_answer_entry.get().strip()
        return question, answers, correct

    def save_question(self):
        question, answers, correct = self.get_question_data()
        if not question or not all(answers) or not all(c in ['1', '2', '3', '4', '5'] for c in correct.replace(" ", "").split(',')):
            messagebox.showerror("오류", "모든 항목을 작성하고 정답은 1~5 범위에서 ,로 구분 입력하세요.")
            return

        try:
            # 텍스트 저장
            with open("exam.txt", "a", encoding="utf-8") as file:
                file.write(f"Question: {question}\n")
                for i, ans in enumerate(answers, 1):
                    file.write(f"{i}) {ans}\n")
                file.write(f"정답 : {correct}\n\n")

            # CSV 저장
            with open("exam.csv", "a", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([question] + answers + [correct])

            # Excel 동기화
            self.update_excel_file()

            self.clear_fields()
            messagebox.showinfo("성공", "문제가 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def update_excel_file(self):
        if os.path.exists("exam.csv"):
            df = pd.read_csv("exam.csv", header=None)
            df.columns = ['Question', '1', '2', '3', '4', '5', 'Answer']
            df.to_excel("exam.xlsx", index=False)

    def show_question_list(self):
        if not os.path.exists("exam.csv"):
            messagebox.showinfo("정보", "저장된 문제가 없습니다.")
            return

        list_win = tk.Toplevel(self.window)
        list_win.title("문제 목록")
        list_win.geometry("700x400")

        tree = ttk.Treeview(list_win, columns=list(range(8)), show="headings")
        columns = ["문제", "1", "2", "3", "4", "5", "정답"]
        for i, col in enumerate(columns):
            tree.heading(i, text=col)
            tree.column(i, width=90)
        tree.pack(fill="both", expand=True)

        # 데이터 로딩
        with open("exam.csv", encoding="utf-8") as f:
            reader = csv.reader(f)
            data = list(reader)

        for row in data:
            tree.insert("", tk.END, values=row)

        # 삭제 버튼
        def delete_selected():
            selected = tree.selection()
            if not selected:
                return
            index = tree.index(selected[0])
            del data[index]
            with open("exam.csv", "w", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(data)
            self.update_excel_file()
            tree.delete(*tree.get_children())
            for row in data:
                tree.insert("", tk.END, values=row)
            messagebox.showinfo("삭제됨", "문제가 삭제되었습니다.")

        tk.Button(list_win, text="선택 삭제", command=delete_selected).pack(pady=5)

    def search_question(self):
        if not os.path.exists("exam.csv"):
            messagebox.showinfo("정보", "저장된 문제가 없습니다.")
            return

        def do_search():
            keyword = entry.get().strip()
            if not keyword:
                return

            results = [row for row in data if keyword in row[0]]
            result_win = tk.Toplevel(search_win)
            result_win.title(f"'{keyword}' 검색 결과")
            result_win.geometry("700x300")

            tree = ttk.Treeview(result_win, columns=list(range(8)), show="headings")
            columns = ["문제", "1", "2", "3", "4", "5", "정답"]
            for i, col in enumerate(columns):
                tree.heading(i, text=col)
                tree.column(i, width=90)
            tree.pack(fill="both", expand=True)

            for row in results:
                tree.insert("", tk.END, values=row)

        with open("exam.csv", encoding="utf-8") as f:
            reader = csv.reader(f)
            data = list(reader)

        search_win = tk.Toplevel(self.window)
        search_win.title("문제 검색")
        search_win.geometry("300x100")
        tk.Label(search_win, text="검색어 입력:").pack()
        entry = tk.Entry(search_win, width=30)
        entry.pack()
        tk.Button(search_win, text="검색", command=do_search).pack(pady=5)

    def start_quiz(self):
        if not os.path.exists("exam.csv"):
            messagebox.showinfo("정보", "저장된 문제가 없습니다.")
            return

        with open("exam.csv", encoding="utf-8") as f:
            questions = list(csv.reader(f))

        if not questions:
            messagebox.showinfo("정보", "문제가 없습니다.")
            return

        q = random.choice(questions)
        q_window = tk.Toplevel(self.window)
        q_window.title("랜덤 퀴즈")
        tk.Label(q_window, text=q[0], wraplength=400, font=("맑은 고딕", 12, "bold")).pack(pady=10)

        var_list = []
        for i in range(1, 6):
            var = tk.IntVar()
            cb = tk.Checkbutton(q_window, text=f"{i}) {q[i]}", variable=var)
            cb.pack(anchor="w")
            var_list.append(var)

        def submit_answer():
            selected = [str(i+1) for i, v in enumerate(var_list) if v.get() == 1]
            correct = [c.strip() for c in q[6].split(',')]
            is_correct = sorted(selected) == sorted(correct)

            # 결과 저장
            with open("exam_results.csv", "a", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([q[0], ",".join(selected), ",".join(correct), "정답" if is_correct else "오답"])

            if is_correct:
                messagebox.showinfo("정답", "정답입니다!")
            else:
                messagebox.showinfo("오답", f"틀렸습니다.\n정답은: {','.join(correct)}")
            q_window.destroy()

        tk.Button(q_window, text="제출", command=submit_answer).pack(pady=10)

    def clear_fields(self):
        self.question_entry.delete("1.0", tk.END)
        for entry in self.options_entries:
            entry.delete(0, tk.END)
        self.correct_answer_entry.delete(0, tk.END)

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    QuizCreator().run()
