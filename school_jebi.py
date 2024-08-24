"""
school_jebi.py    제비뽑기    Ver 1.0_240602
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import random
import csv

# 발표 순서를 랜덤하게 생성하고 결과를 표시하는 함수
def generate_order():
    global random_order  # 저장 기능에서 사용하기 위해 전역 변수로 선언
    subject = subject_entry.get()  # 과목명 가져오기
    if not subject.strip():
        messagebox.showwarning("경고", "과목명을 입력해주세요.")
        return
    N = int(students_entry.get())  # 사용자 입력을 받아 학생 수 N을 설정
    random_order = random.sample(range(1, N + 1), N)  # 1부터 N까지 숫자 중 N개를 랜덤하게 선택

    # 첫 번째 발표할 학생의 번호를 팝업창으로 보여주기
    messagebox.showinfo("첫 번째 발표자", f"첫 번째 발표할 학생의 번호는 {random_order[0]}입니다.")

    # 표(treeview)를 초기화
    for i in tree.get_children():
        tree.delete(i)

    # 랜덤하게 생성된 발표 순서를 표에 추가
    for i, number in enumerate(random_order, start=1):
        tree.insert("", "end", values=(i, number))

# 발표 순서를 CSV 파일로 저장하는 함수
def save_to_csv():
    if random_order:
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV 파일", "*.csv")])
        if filename:
            with open(filename, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["발표 순서", "학생 번호"])
                for i, number in enumerate(random_order, start=1):
                    writer.writerow([i, number])
            messagebox.showinfo("저장 성공", f"'{filename}'로 저장되었습니다.")
    else:
        messagebox.showwarning("경고", "먼저 발표 순서를 생성해주세요.")

# 발표 순서를 TXT 파일로 저장하는 함수
def save_to_txt():
    if random_order:
        filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("텍스트 파일", "*.txt")])
        if filename:
            with open(filename, mode="w", encoding="utf-8") as file:
                for i, number in enumerate(random_order, start=1):
                    file.write(f"발표 순서 {i}: 학생 번호 {number}\n")
            messagebox.showinfo("저장 성공", f"'{filename}'로 저장되었습니다.")
    else:
        messagebox.showwarning("경고", "먼저 발표 순서를 생성해주세요.")

# 간단한 인쇄 메시지를 표시하는 함수
def print_order():
    if random_order:
        messagebox.showinfo("인쇄", "인쇄 기능은 이 예제에서 구현되지 않았습니다.")
    else:
        messagebox.showwarning("경고", "먼저 발표 순서를 생성해주세요.")

# 메뉴에서 호출할 기능들
def menu_save_to_csv():
    save_to_csv()

def menu_save_to_txt():
    save_to_txt()

def menu_print_order():
    print_order()

# 기본 tkinter 윈도우 설정
window = tk.Tk()
window.title("발표 순서 생성기")

# 전역 변수 초기화
random_order = []

# 메뉴 바 설정
menubar = tk.Menu(window)
window.config(menu=menubar)

# 저장 메뉴 추가
save_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="저장", menu=save_menu)
save_menu.add_command(label="CSV로 저장", command=menu_save_to_csv)
save_menu.add_command(label="TXT로 저장", command=menu_save_to_txt)
save_menu.add_separator()
save_menu.add_command(label="인쇄", command=menu_print_order)

# 사용자 입력을 받는 프레임 설정
input_frame = tk.Frame(window)
input_frame.pack(pady=10)

# 과목명 입력
subject_label = tk.Label(input_frame, text="과목명:")
subject_label.pack(side=tk.LEFT)
subject_entry = tk.Entry(input_frame)
subject_entry.pack(side=tk.LEFT, padx=5)

# 학생 수 입력
students_label = tk.Label(input_frame, text="학생 수:")
students_label.pack(side=tk.LEFT)
students_entry = tk.Entry(input_frame, justify='center')
students_entry.pack(side=tk.LEFT, padx=5)

# 발표 순서를 생성하는 버튼 설정
generate_button = tk.Button(input_frame, text="발표 순서 생성", bg="green", fg="white", command=generate_order)
generate_button.pack(side=tk.LEFT, padx=5)

# 표시할 발표 순서를 나타내는 표(treeview) 설정
tree_frame = tk.Frame(window)
tree_frame.pack(pady=10)

tree = ttk.Treeview(tree_frame, columns=("발표 순서", "학생 번호"), show="headings")
tree.heading("발표 순서", text="발표 순서")
tree.heading("학생 번호", text="학생 번호")
tree.column("발표 순서", anchor='center')
tree.column("학생 번호", anchor='center')

# 스크롤 바 추가
scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
scrollbar.pack(side='right', fill='y')
tree.configure(yscrollcommand=scrollbar.set)
tree.pack()

window.mainloop()
