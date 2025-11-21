"""
exam_seat.py   시험 좌석 배정  Ver 1.1_251116
"""
import tkinter as tk
from tkinter import messagebox, filedialog
import csv
import os
import random
from datetime import datetime

# 설정
NUM_COLS = 10
NUM_ROWS = 5
PAIR_SIZE = 2
PAIR_COUNT = NUM_COLS // PAIR_SIZE
TOTAL_SLOTS = NUM_COLS * NUM_ROWS
CSV_PATH = "seats.csv"

# 전역 변수
selected_states = [False] * TOTAL_SLOTS
time_label = None
buttons = []
counter_label = None


def read_csv_labels(path):
    """CSV 파일에서 좌석 정보를 읽어옴"""
    labels = []
    if os.path.exists(path):
        try:
            with open(path, newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    for cell in row:
                        text = cell.strip()
                        if text.startswith('*'):
                            labels.append(text[1:])
                            idx = len(labels) - 1
                            selected_states[idx] = True
                        else:
                            labels.append(text)
        except Exception as e:
            messagebox.showerror("오류", f"CSV 파일 읽기 실패:\n{e}")

    # 빈 슬롯 채우기
    if len(labels) < TOTAL_SLOTS:
        labels += [''] * (TOTAL_SLOTS - len(labels))
    else:
        labels = labels[:TOTAL_SLOTS]
    return labels


def save_selected_states(labels):
    """선택 상태를 CSV 파일에 저장"""
    try:
        with open(CSV_PATH, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            row = []
            for i in range(TOTAL_SLOTS):
                label = labels[i]
                if selected_states[i] and label:
                    row.append(f"*{label}")
                else:
                    row.append(label)
            writer.writerow(row)
    except Exception as e:
        messagebox.showerror("저장 오류", f"파일 저장 실패:\n{e}")


def update_counter():
    """선택된 좌석 수 업데이트"""
    if counter_label:
        count = sum(1 for i in range(TOTAL_SLOTS) if selected_states[i])
        counter_label.config(text=f"선택된 좌석: {count} / {TOTAL_SLOTS}")


def on_button_click(btn, text, index, labels):
    """버튼 클릭 시 선택 상태 토글"""
    selected_states[index] = not selected_states[index]

    # 색상 업데이트
    if selected_states[index]:
        btn.config(bg='#b0e0e6', fg='black')  # 선택됨
    else:
        btn.config(bg='#dcdcdc', fg='#dcdcdc')  # 선택 안됨

    save_selected_states(labels)
    update_counter()


def select_all_seats(labels):
    """모든 좌석 선택"""
    for i in range(TOTAL_SLOTS):
        if labels[i]:  # 빈 좌석이 아닌 경우만
            selected_states[i] = True
            if buttons[i]:
                buttons[i].config(bg='#b0e0e6', fg='black')
    save_selected_states(labels)
    update_counter()


def deselect_all_seats(labels):
    """모든 좌석 선택 해제"""
    for i in range(TOTAL_SLOTS):
        selected_states[i] = False
        if buttons[i] and labels[i]:
            buttons[i].config(bg='#dcdcdc', fg='#dcdcdc')
    save_selected_states(labels)
    update_counter()


def generate_seat_numbers(labels):
    """좌석 번호 자동 생성 (1-1, 1-2, ..., 5-10)"""
    if messagebox.askyesno("좌석 번호 생성",
                           "현재 좌석 정보를 모두 지우고\n좌석 번호(1 ~ 50)로 채우시겠습니까?"):
        n=0
        for r in range(NUM_ROWS):
            for c in range(NUM_COLS):
                idx = r * NUM_COLS + c
                n = n + 1  # 번호
                labels[idx] = f"{n}"
                # labels[idx] = f"{r + 1}-{c + 1}"
                selected_states[idx] = False

        update_buttons(labels)
        save_selected_states(labels)
        update_counter()
        messagebox.showinfo("완료", "좌석 번호가 생성되었습니다.")


def load_student_list(labels):
    """출석부(CSV/TXT) 불러오기"""
    file_path = filedialog.askopenfilename(
        title="출석부 파일 선택",
        filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")]
    )

    if not file_path:
        return

    try:
        students = []
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            if file_path.endswith('.csv'):
                reader = csv.reader(f)
                for row in reader:
                    for cell in row:
                        name = cell.strip()
                        if name:
                            students.append(name)
            else:  # TXT 파일
                for line in f:
                    name = line.strip()
                    if name:
                        students.append(name)

        if not students:
            messagebox.showwarning("경고", "파일에서 이름을 찾을 수 없습니다.")
            return

        if len(students) > TOTAL_SLOTS:
            messagebox.showwarning("경고",
                                   f"학생 수({len(students)})가 좌석 수({TOTAL_SLOTS})를 초과합니다.\n"
                                   f"처음 {TOTAL_SLOTS}명만 배치됩니다.")
            students = students[:TOTAL_SLOTS]

        # 순서대로 배치
        for i in range(len(students)):
            labels[i] = students[i]
            selected_states[i] = False

        # 나머지는 빈 좌석
        for i in range(len(students), TOTAL_SLOTS):
            labels[i] = ''
            selected_states[i] = False

        update_buttons(labels)
        save_selected_states(labels)
        update_counter()
        messagebox.showinfo("완료", f"{len(students)}명의 학생이 배치되었습니다.")

    except Exception as e:
        messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")


def random_assign_seats(labels):
    """랜덤 좌석 배치"""
    # 빈 좌석이 아닌 이름들만 추출
    names = [labels[i] for i in range(TOTAL_SLOTS) if labels[i]]

    if not names:
        messagebox.showwarning("경고", "배치할 좌석 정보가 없습니다.")
        return

    if messagebox.askyesno("랜덤 배치",
                           f"{len(names)}개의 좌석을 무작위로 섞으시겠습니까?"):
        # 이름들을 랜덤하게 섞음
        random.shuffle(names)

        # 섞인 이름들을 좌석에 배치
        name_idx = 0
        for i in range(TOTAL_SLOTS):
            if name_idx < len(names):
                labels[i] = names[name_idx]
                name_idx += 1
            else:
                labels[i] = ''
            selected_states[i] = False

        update_buttons(labels)
        save_selected_states(labels)
        update_counter()
        messagebox.showinfo("완료", "랜덤 배치가 완료되었습니다.")


def update_buttons(labels):
    """모든 버튼 업데이트"""
    for i in range(TOTAL_SLOTS):
        if buttons[i]:
            if labels[i] == '':
                buttons[i].config(text='', state='disabled', bg='#dcdcdc')
            else:
                buttons[i].config(text=labels[i], state='normal')
                if selected_states[i]:
                    buttons[i].config(bg='#b0e0e6', fg='black')
                else:
                    buttons[i].config(bg='#dcdcdc', fg='black')


def update_time(label):
    """실시간 시계 업데이트"""
    now = datetime.now().strftime("%H:%M:%S")
    label.config(text=f"현재 시간: {now}")
    label.after(1000, update_time, label)


def toggle_time_label_color(event):
    """'t' 키로 시계 표시/숨김 토글"""
    current_fg = time_label.cget("fg")
    if current_fg == "#f0f0f0":
        time_label.config(fg="black")
    else:
        time_label.config(fg="#f0f0f0")


def open_edit_window(labels):
    """좌석 정보 수정 창 열기"""
    edit_win = tk.Toplevel()
    edit_win.title("좌석 정보 수정")
    edit_win.geometry("700x300")

    tk.Label(edit_win, text="좌석 정보 수정 (5행 × 10열)",
             font=("맑은 고딕", 16, "bold"), pady=10).pack()

    frame = tk.Frame(edit_win)
    frame.pack(padx=20, pady=10)

    # 입력 필드 생성
    entries = []
    for r in range(NUM_ROWS):
        row_entries = []
        for c in range(NUM_COLS):
            idx = r * NUM_COLS + c
            entry = tk.Entry(frame, width=8, font=("맑은 고딕", 10))
            entry.grid(row=r, column=c, padx=2, pady=2)
            entry.insert(0, labels[idx])
            row_entries.append(entry)
        entries.append(row_entries)

    def save_changes():
        """수정 사항 저장"""
        new_labels = []
        for r in range(NUM_ROWS):
            for c in range(NUM_COLS):
                new_labels.append(entries[r][c].get().strip())

        # CSV 저장
        try:
            with open(CSV_PATH, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                row = []
                for i in range(TOTAL_SLOTS):
                    label = new_labels[i]
                    if selected_states[i] and label:
                        row.append(f"*{label}")
                    else:
                        row.append(label)
                writer.writerow(row)

            # 버튼 업데이트
            for i in range(TOTAL_SLOTS):
                labels[i] = new_labels[i]

            update_buttons(labels)
            update_counter()

            messagebox.showinfo("저장 완료", "좌석 정보가 저장되었습니다.")
            edit_win.destroy()
        except Exception as e:
            messagebox.showerror("저장 실패", f"저장 중 오류 발생:\n{e}")

    # 버튼
    btn_frame = tk.Frame(edit_win)
    btn_frame.pack(pady=20)

    tk.Button(btn_frame, text="저장", font=("맑은 고딕", 12, "bold"),
              width=10, command=save_changes, bg='#4CAF50', fg='white').pack(side='left', padx=5)
    tk.Button(btn_frame, text="취소", font=("맑은 고딕", 12),
              width=10, command=edit_win.destroy, bg='#f44336', fg='white').pack(side='left', padx=5)


def make_buttons_from_labels(root, labels):
    """메인 UI 생성"""
    global time_label, buttons, counter_label

    # 타이틀
    title_label = tk.Label(root, text="[       시험 좌석 위치 비번호 (칠판)       ]",
                           font=("맑은 고딕", 42, "bold"), pady=10, bg="silver")
    title_label.pack()

    # 시계
    time_label = tk.Label(root, font=("맑은 고딕", 72, "bold"),
                          pady=5, bg="#f0f0f0", fg="#f0f0f0")
    time_label.pack()
    update_time(time_label)

    # 't' 키 바인딩
    root.bind("<t>", toggle_time_label_color)
    root.bind("<T>", toggle_time_label_color)

    # 좌석 배치 프레임
    main_frame = tk.Frame(root, padx=60, pady=16)
    main_frame.pack()

    buttons = [None] * TOTAL_SLOTS

    # 좌석 버튼 생성
    for r in range(NUM_ROWS):
        for pair in range(PAIR_COUNT):
            pair_frame = tk.Frame(main_frame)
            pair_frame.grid(row=r, column=pair, padx=(0, 40), pady=18)

            for i in range(PAIR_SIZE):
                col_index = pair * PAIR_SIZE + i
                slot_index = r * NUM_COLS + col_index
                label = labels[slot_index]

                btn_font = ("맑은 고딕", 24, "bold")

                if label == '':
                    # 빈 좌석
                    btn = tk.Button(pair_frame, text='', width=8, height=1, state='disabled',
                                    font=btn_font, border=2, bg='#dcdcdc')
                else:
                    # 좌석 있음
                    if selected_states[slot_index]:
                        btn_bg = '#b0e0e6'  # 선택됨
                        btn_fg = 'black'
                    else:
                        btn_bg = '#dcdcdc'  # 선택 안됨
                        btn_fg = 'black'

                    btn = tk.Button(
                        pair_frame, text=label, width=8,
                        height=1, font=btn_font, bg=btn_bg, fg=btn_fg, border=2
                    )
                    btn.config(command=lambda b=btn, txt=label, idx=slot_index:
                    on_button_click(b, txt, idx, labels))

                btn.pack(side='left', padx=(0, 2))
                buttons[slot_index] = btn

    # 하단 버튼 프레임
    bottom_frame = tk.Frame(root)
    bottom_frame.pack(pady=20)

    # 첫 번째 줄 버튼
    row1_frame = tk.Frame(bottom_frame)
    row1_frame.pack(pady=5)

    tk.Button(row1_frame, text="전체 선택", font=("맑은 고딕", 10),
              width=15, height=1, bg='#dcdcdc', fg='blue',
              command=lambda: select_all_seats(labels)).pack(side='left', padx=5)

    tk.Button(row1_frame, text="전체 해제", font=("맑은 고딕", 10),
              width=15, height=1, bg='#dcdcdc', fg='blue',
              command=lambda: deselect_all_seats(labels)).pack(side='left', padx=5)

    tk.Button(row1_frame, text="수정", font=("맑은 고딕", 10),
              width=15, height=1, bg='#dcdcdc', fg='blue',
              command=lambda: open_edit_window(labels)).pack(side='left', padx=5)

    tk.Button(row1_frame, text="좌석 번호 생성", font=("맑은 고딕", 10),
              width=15, height=1, bg='#dcdcdc', fg='blue',
              command=lambda: generate_seat_numbers(labels)).pack(side='left', padx=5)

    tk.Button(row1_frame, text="출석부 불러오기", font=("맑은 고딕", 10),
              width=15, height=1, bg='#dcdcdc', fg='blue',
              command=lambda: load_student_list(labels)).pack(side='left', padx=5)

    tk.Button(row1_frame, text="랜덤 배치", font=("맑은 고딕", 10),
              width=15, height=1, bg='#dcdcdc', fg='blue',
              command=lambda: random_assign_seats(labels)).pack(side='left', padx=5)

    # 카운터 (화면 가장 아래)
    counter_label = tk.Label(root, text="", font=("맑은 고딕", 12), pady=10)
    counter_label.pack(side='bottom')
    update_counter()


if __name__ == "__main__":
    root = tk.Tk()
    root.title("시험 좌석 배치")
    root.state('zoomed')

    try:
        labels = read_csv_labels(CSV_PATH)
        make_buttons_from_labels(root, labels)
    except Exception as e:
        messagebox.showerror("오류", f"프로그램 시작 중 오류 발생:\n{e}")

    root.mainloop()