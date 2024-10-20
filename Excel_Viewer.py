"""
Excel_Viewer.py  엑셀 tkinter view Ver 1.0_241021
"""
from tkinter import *
from tkinter import messagebox
from tkinter import ttk, filedialog
import pandas as pd

root = Tk()
root.title("Excel Viewer")
root.geometry('1000x400+200+200')

# 엑셀 파일 열기 함수
def Open_file():
    filename = filedialog.askopenfilename(title="Open a File", filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*")))

    if filename:
        try:
            df = pd.read_excel(filename)
            df.fillna("", inplace=True)  # NaN 및 빈 셀을 공백으로 대체
        except Exception as e:
            messagebox.showerror("오류", "파일을 열 수 없습니다: {}".format(str(e)))
            return

        # Treeview 초기화
        tree.delete(*tree.get_children())

        # 열 제목 설정
        tree['columns'] = list(df.columns)
        tree['show'] = "headings"

        for col in tree['columns']:
            tree.heading(col, text=col)
            tree.column(col, width=100)  # 열 너비 기본값 설정

        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            tree.insert("", "end", values=row)

        # 열 너비 자동 조정
        for col in tree['columns']:
            tree.column(col, width=150, minwidth=100)

# 프로그램 종료 함수
def Exit():
    if messagebox.askyesno("Exit", "정말 종료하시겠습니까?"):
        root.destroy()

# Help 메뉴에서 호출할 팝업 함수
def show_help():
    help_content=("tkinter 모듈을 사용한 엑셀 뷰어 입니다.\n버튼을 없애고 메뉴로 대체 했습니다."+
                  "\n스크롤바를 추가하고 공백은 빈칸으로 출력됩니다.")
    messagebox.showinfo("Help", help_content)

# 아이콘 설정
image_icon = PhotoImage(file="excel.png")  # 파일 경로에 맞게 수정 필요
root.iconphoto(False, image_icon)

# 프레임 설정
frame = Frame(root)
frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")  # 좌우 10px 여백 유지

# Treeview 및 스크롤바
tree_frame = Frame(frame)
tree_frame.grid(row=0, column=0, sticky="nsew")

# 수직 스크롤바
tree_scroll_y = Scrollbar(tree_frame)
tree_scroll_y.pack(side=RIGHT, fill=Y)

# 수평 스크롤바
tree_scroll_x = Scrollbar(tree_frame, orient='horizontal')
tree_scroll_x.pack(side=BOTTOM, fill=X)

# Treeview
tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
tree.pack(fill=BOTH, expand=True)

# 스크롤바 설정
tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

# 메뉴 설정
my_menu = Menu(root)
root.config(menu=my_menu)

# 파일 메뉴 추가
file_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open", command=Open_file)
file_menu.add_command(label="Exit", command=Exit)

# Help 메뉴 추가
help_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="Help", menu=help_menu)
help_menu.add_command(label="About", command=show_help)

# 창 크기에 맞게 자동 확장 설정
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
frame.grid_rowconfigure(0, weight=1)
frame.grid_columnconfigure(0, weight=1)

root.mainloop()
