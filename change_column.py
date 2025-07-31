"""
change_column.py 열 순서 변경  Ver 1.1_250801
"""
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font
import os


class DragDropListbox(tk.Listbox):
    """드래그 앤 드롭이 가능한 리스트박스"""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.bind('<Button-1>', self.on_click)
        self.bind('<B1-Motion>', self.on_drag)
        self.bind('<ButtonRelease-1>', self.on_drop)
        self.drag_start_index = None

    def on_click(self, event):
        self.drag_start_index = self.nearest(event.y)

    def on_drag(self, event):
        if self.drag_start_index is not None:
            current_index = self.nearest(event.y)
            if current_index != self.drag_start_index:
                # 시각적 피드백을 위한 선택 변경
                self.selection_clear(0, tk.END)
                self.selection_set(current_index)

    def on_drop(self, event):
        if self.drag_start_index is not None:
            drop_index = self.nearest(event.y)
            if drop_index != self.drag_start_index:
                # 아이템 이동
                item = self.get(self.drag_start_index)
                self.delete(self.drag_start_index)
                self.insert(drop_index, item)
                self.selection_clear(0, tk.END)
                self.selection_set(drop_index)
        self.drag_start_index = None


class ExcelColumnReorderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 열 순서 변경 도구")
        self.root.geometry("600x700")
        self.root.resizable(True, True)

        self.setup_styles()
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.current_columns = []

        self.create_widgets()

        # 프로그램 시작 시 필수 라이브러리 확인
        self.check_prerequisites()

    def setup_styles(self):
        """UI 스타일 설정"""
        self.title_font = font.Font(family="D2Coding", size=14, weight="bold")
        self.button_font = font.Font(family="D2Coding", size=12)
        self.listbox_font = font.Font(family="D2Coding", size=11)

    def check_prerequisites(self):
        """Pandas Excel 엔진 라이브러리 설치 여부 확인"""
        try:
            import openpyxl
            import xlrd
        except ImportError:
            messagebox.showwarning(
                "필수 라이브러리 설치 필요",
                "이 프로그램은 엑셀 파일을 처리하기 위해 'openpyxl'과 'xlrd' 라이브러리가 필요합니다.\n\n"
                "터미널(또는 명령 프롬프트)에서 아래 명령어를 실행하여 설치해주세요:\n"
                "pip install openpyxl xlrd"
            )

    def create_widgets(self):
        # 메인 컨테이너
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # 파일 선택 섹션
        self.create_file_section(main_frame)
        # 열 순서 편집 섹션
        self.create_column_edit_section(main_frame)
        # 제어 버튼 섹션
        self.create_control_buttons(main_frame)
        # 실행 버튼
        self.create_execute_button(main_frame)

    def create_file_section(self, parent):
        """파일 선택 섹션 생성"""
        file_frame = tk.LabelFrame(parent, text="파일 선택", font=self.title_font, padx=10, pady=10)
        file_frame.pack(fill="x", pady=(0, 10))

        # 입력 파일
        input_frame = tk.Frame(file_frame)
        input_frame.pack(fill="x", pady=5)
        tk.Label(input_frame, text="원본 파일:", width=10, anchor="w", font=("D2Coding", 12)).pack(side="left")
        tk.Entry(input_frame, textvariable=self.input_file_path, state="readonly", font=("D2Coding", 12)).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(input_frame, text="찾아보기", command=self.browse_input_file, width=12,
                  bg="#2196F3", fg="white", font=self.button_font).pack(side="right")

        # 출력 파일
        output_frame = tk.Frame(file_frame)
        output_frame.pack(fill="x", pady=5)
        tk.Label(output_frame, text="저장 파일:", width=10, anchor="w", font=("D2Coding", 12)).pack(side="left")
        tk.Entry(output_frame, textvariable=self.output_file_path, state="readonly", font=("D2Coding", 12)).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(output_frame, text="저장경로", command=self.browse_output_file, width=12,
                  bg="#2196F3", fg="white", font=self.button_font).pack(side="right")

    def create_column_edit_section(self, parent):
        """열 순서 편집 섹션 생성"""
        edit_frame = tk.LabelFrame(parent, text="열 순서 편집", font=self.title_font, padx=10, pady=10)
        edit_frame.pack(fill="both", expand=True, pady=(0, 10))

        # 안내 문구
        info_label = tk.Label(edit_frame,
                              text="• 드래그 앤 드롭으로 순서 변경 • 화살표 버튼으로 미세 조정",
                              fg="blue", font=("D2Coding", 11))
        info_label.pack(pady=(0, 10))

        # 메인 편집 영역
        main_edit_frame = tk.Frame(edit_frame)
        main_edit_frame.pack(fill="both", expand=True)

        # 왼쪽: 현재 열 순서
        left_frame = tk.Frame(main_edit_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        tk.Label(left_frame, text="현재 열 순서", font=self.title_font).pack()

        listbox_frame = tk.Frame(left_frame)
        listbox_frame.pack(fill="both", expand=True)

        self.column_listbox = DragDropListbox(listbox_frame, height=15, selectmode=tk.SINGLE, font=self.listbox_font)
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical")
        self.column_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.column_listbox.yview)

        self.column_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 오른쪽: 제어 버튼들
        right_frame = tk.Frame(main_edit_frame)
        right_frame.pack(side="right", fill="y", padx=(5, 0))

        self.create_move_buttons(right_frame)

    def create_move_buttons(self, parent):
        """이동 버튼들 생성"""
        tk.Label(parent, text="순서 조정", font=self.title_font).pack(pady=(0, 10))

        button_config = {"width": 12, "font": self.button_font, "bg": "#4CAF50", "fg": "white"}

        tk.Button(parent, text="▲ 맨 위로", command=self.move_to_top, **button_config).pack(pady=2)
        tk.Button(parent, text="↑ 위로", command=self.move_up, **button_config).pack(pady=2)
        tk.Button(parent, text="↓ 아래로", command=self.move_down, **button_config).pack(pady=2)
        tk.Button(parent, text="▼ 맨 아래로", command=self.move_to_bottom, **button_config).pack(pady=2)

        tk.Frame(parent, height=20).pack()

        # 선택 관련 버튼
        tk.Label(parent, text="선택", font=self.title_font).pack(pady=(0, 5))
        tk.Button(parent, text="모두 선택", command=self.select_all,
                  bg="#FF9800", fg="white", **{"width": 12, "font": self.button_font}).pack(pady=2)
        tk.Button(parent, text="선택 해제", command=self.deselect_all,
                  bg="#FF9800", fg="white", **{"width": 12, "font": self.button_font}).pack(pady=2)

    def create_control_buttons(self, parent):
        """제어 버튼 섹션 생성"""
        control_frame = tk.Frame(parent)
        control_frame.pack(fill="x", pady=(0, 10))
        button_config = {"font": self.button_font, "padx": 20}

        tk.Button(control_frame, text="열 정보 새로고침", command=self.refresh_columns,
                  bg="#9C27B0", fg="white", **button_config).pack(side="left", padx=5)
        tk.Button(control_frame, text="순서 초기화", command=self.reset_order,
                  bg="#FF5722", fg="white", **button_config).pack(side="left", padx=5)
        tk.Button(control_frame, text="미리보기", command=self.preview_result,
                  bg="#607D8B", fg="white", **button_config).pack(side="right", padx=5)

    def create_execute_button(self, parent):
        """실행 버튼 생성"""
        execute_frame = tk.Frame(parent)
        execute_frame.pack(fill="x")
        self.execute_button = tk.Button(execute_frame, text="열 순서 변경하여 저장",
                                        command=self.execute_reorder,
                                        height=1, font=("D2Coding", 14, "bold"),
                                        bg="#4CAF50", fg="white")
        self.execute_button.pack(fill="x")

    def browse_input_file(self):
        """입력 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_file_path.set(file_path)

            # os.path.splitext를 사용하여 파일명과 확장자 분리
            base_name, ext = os.path.splitext(file_path)

            # 새로운 파일명 생성
            output_path = f"{base_name}_reordered{ext}"
            self.output_file_path.set(output_path)

            self.load_columns()

    def browse_output_file(self):
        """출력 파일 경로 선택"""
        file_path = filedialog.asksaveasfilename(
            title="저장할 파일명 지정",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.output_file_path.set(file_path)

    def load_columns(self):
        """엑셀 파일의 열 정보를 로드"""
        input_path = self.input_file_path.get()
        if not input_path:
            return

        try:
            # 파일 확장자를 기반으로 pandas 엔진 선택
            if input_path.endswith('.xlsx'):
                df = pd.read_excel(input_path, engine='openpyxl')
            elif input_path.endswith('.xls'):
                df = pd.read_excel(input_path, engine='xlrd')
            else:
                messagebox.showerror("오류", "지원되지 않는 파일 형식입니다. .xlsx 또는 .xls 파일을 선택해주세요.")
                return

            self.current_columns = df.columns.tolist()
            self.refresh_listbox()
            messagebox.showinfo("성공", f"총 {len(self.current_columns)}개의 열을 로드했습니다.")
        except FileNotFoundError:
            messagebox.showerror("오류", "파일을 찾을 수 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")

    def refresh_listbox(self):
        """리스트박스 새로고침"""
        self.column_listbox.delete(0, tk.END)
        for i, col in enumerate(self.current_columns):
            self.column_listbox.insert(tk.END, f"{i + 1:2d}. {col}")

    def refresh_columns(self):
        """열 정보 새로고침"""
        self.load_columns()

    def reset_order(self):
        """순서 초기화"""
        if messagebox.askyesno("확인", "열 순서를 원래대로 초기화하시겠습니까?"):
            self.load_columns()

    def get_selected_index(self):
        """선택된 항목의 인덱스 반환"""
        selection = self.column_listbox.curselection()
        return selection[0] if selection else None

    def move_to_top(self):
        """선택된 항목을 맨 위로"""
        index = self.get_selected_index()
        if index is not None and index > 0:
            item = self.current_columns.pop(index)
            self.current_columns.insert(0, item)
            self.refresh_listbox()
            self.column_listbox.selection_set(0)

    def move_up(self):
        """선택된 항목을 위로"""
        index = self.get_selected_index()
        if index is not None and index > 0:
            self.current_columns[index], self.current_columns[index - 1] = \
                self.current_columns[index - 1], self.current_columns[index]
            self.refresh_listbox()
            self.column_listbox.selection_set(index - 1)

    def move_down(self):
        """선택된 항목을 아래로"""
        index = self.get_selected_index()
        if index is not None and index < len(self.current_columns) - 1:
            self.current_columns[index], self.current_columns[index + 1] = \
                self.current_columns[index + 1], self.current_columns[index]
            self.refresh_listbox()
            self.column_listbox.selection_set(index + 1)

    def move_to_bottom(self):
        """선택된 항목을 맨 아래로"""
        index = self.get_selected_index()
        if index is not None and index < len(self.current_columns) - 1:
            item = self.current_columns.pop(index)
            self.current_columns.append(item)
            self.refresh_listbox()
            self.column_listbox.selection_set(len(self.current_columns) - 1)

    def select_all(self):
        """모든 항목 선택"""
        self.column_listbox.selection_set(0, tk.END)

    def deselect_all(self):
        """모든 선택 해제"""
        self.column_listbox.selection_clear(0, tk.END)

    def preview_result(self):
        """결과 미리보기"""
        if not self.current_columns:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 선택하고 열 정보를 로드해주세요.")
            return
        preview_text = "변경된 열 순서:\n\n"
        for i, col in enumerate(self.current_columns):
            preview_text += f"{i + 1:2d}. {col}\n"

        preview_window = tk.Toplevel(self.root)
        preview_window.title("순서 변경 미리보기")
        preview_window.geometry("400x500")
        preview_window.resizable(False, False)
        text_widget = tk.Text(preview_window, wrap=tk.WORD, padx=10, pady=10)
        scrollbar = tk.Scrollbar(preview_window, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        text_widget.insert("1.0", preview_text)
        text_widget.config(state="disabled")

    def execute_reorder(self):
        """열 순서 변경 실행"""
        input_path = self.input_file_path.get()
        output_path = self.output_file_path.get()

        if not input_path:
            messagebox.showwarning("입력 오류", "원본 엑셀 파일을 선택해주세요.")
            return
        if not output_path:
            messagebox.showwarning("입력 오류", "저장할 파일 경로를 지정해주세요.")
            return
        if not self.current_columns:
            messagebox.showwarning("입력 오류", "열 정보를 먼저 로드해주세요.")
            return

        try:
            # 엑셀 파일 읽기
            if input_path.endswith('.xlsx'):
                df = pd.read_excel(input_path, engine='openpyxl')
            elif input_path.endswith('.xls'):
                df = pd.read_excel(input_path, engine='xlrd')
            else:
                messagebox.showerror("오류", "지원되지 않는 파일 형식입니다. .xlsx 또는 .xls 파일을 선택해주세요.")
                return

            # 열 순서 변경
            df_reordered = df[self.current_columns]

            # 저장
            df_reordered.to_excel(output_path, index=False)
            messagebox.showinfo("완료",
                                f"열 순서가 성공적으로 변경되어 저장되었습니다!\n\n"
                                f"저장 위치: {output_path}\n"
                                f"총 열 개수: {len(self.current_columns)}개")
        except FileNotFoundError:
            messagebox.showerror("오류", "지정된 파일을 찾을 수 없습니다. 경로를 확인해주세요.")
        except KeyError as e:
            messagebox.showerror("오류", f"선택한 열 중 일부가 원본 파일에 없습니다:\n{e}")
        except Exception as e:
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelColumnReorderApp(root)
    try:
        root.iconbitmap('codemy.ico')
    except tk.TclError:
        pass
    root.mainloop()