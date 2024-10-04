"""
school_student.py Tkinter Frame Ver 1.0_241005
"""
import os
from tkinter import *
from tkinter import ttk
from PIL import Image,ImageTk
from tkinter import messagebox
import openpyxl

# Color, font
bl_fg='crimson'
lf_bg = 'LightSkyBlue'
rtf_bg = 'DeepSkyBlue'
btn_hlb_bg = 'darkturquoise'
lo_bg='forestgreen'
tit_font = ('Georgia', 36, 'bold')
entry_font = ('맑은 고딕', 12)
btn_font = ('맑은 고딕', 14, 'bold')

class student:
    def __init__(self,root):
        self.excel_file = 'StudentDatas.xlsx'
        self.root=root
        self.root.geometry("1530x880+10+10")
        self.root.title("학생 관리 시스템")

        # Variables
        self.var_numb=StringVar()        # 학번
        self.var_name=StringVar()       # 성명
        self.var_gender=StringVar()     # 성별
        self.var_regist=StringVar()     # 등록
        self.var_phone=StringVar()      # 전화번호
        self.var_middle=StringVar()     # 출신중학교
        self.var_jumin1=StringVar()     # 주민번호1
        self.var_jumin2=StringVar()     # 주민번호1
        self.var_address=StringVar()    # 주소
        self.var_post=StringVar()       # 우편번호
        self.var_email=StringVar()      # 이메일
        self.var_fatname=StringVar()    # 부성명
        self.var_fatphone=StringVar()   # 부전화
        self.var_motname=StringVar()    # 모성명
        self.var_motphone=StringVar()   # 모전화
        self.var_club=StringVar()       # 특활부서
        self.var_religion=StringVar()   # 종교
        self.var_hobby=StringVar()      # 특별활동

        self.is_first_input = True
        self.search_by = StringVar()
        self.search_txt = StringVar()

        self.photo_path='images/'
        self.default_img_path='./images/default_photo.jpg'

        # Create Frames and Widgets
        self.create_widgets()

        self.load_excel()
        self.fetch_data()

        try:
            self.load_excel()
            self.fetch_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel: {e}")

    def create_widgets(self):
        # Title Frame
        lbl_title = Label(self.root, text='STUDENTS MANAGEMENT SYSTEM', font=tit_font, fg=bl_fg, bg=lf_bg)
        lbl_title.place(x=0, y=0, width=1530, height=100)

        # logo
        logo=Image.open('images/yushin_logo.jpg')
        logo=logo.resize((80,80),Image.LANCZOS)
        self.photo=ImageTk.PhotoImage(logo)
        self.img_1=Label(self.root,image=self.photo)
        self.img_1.place(x=30,y=5,width=80,height=80)

        # Image Logo Frame
        img_frame = Frame(self.root, relief=RIDGE)
        img_frame.place(x=0, y=100, width=1530, height=120)

        # 1st
        img_1=Image.open('images/emp5.jpg')
        img_1=img_1.resize((540,120),Image.LANCZOS)
        self.photo1=ImageTk.PhotoImage(img_1)
        self.img_1=Label(img_frame,image=self.photo1)
        self.img_1.place(x=0,y=0,width=540,height=120)

        # 2st
        img_2=Image.open('images/emp2.jpg')
        img_2=img_2.resize((500,160),Image.LANCZOS)
        self.photo2=ImageTk.PhotoImage(img_2)
        self.img_2=Label(img_frame,image=self.photo2)
        self.img_2.place(x=520,y=0,width=520,height=120)

        # 3st
        img_3=Image.open('images/yushin.jpg')
        img_3=img_3.resize((500,120),Image.LANCZOS)
        self.photo3=ImageTk.PhotoImage(img_3)
        self.img_3=Label(img_frame,image=self.photo3)
        self.img_3.place(x=1040,y=0,width=500,height=120)

        # Main Frame
        Main_frame = Frame(self.root, relief=RIDGE,bg='white')
        Main_frame.place(x=0, y=220, width=1530, height=580)

        # 학생 정보 Upper frame
        upper_frame = LabelFrame(Main_frame,bd=0,relief=RIDGE, bg='white', font=('맑은 고딕', 11, 'bold'), fg='black')
        upper_frame.place(x=0, y=10, width=1280, height=300)

        # label : Num 학번
        lal_Numb=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="학번:",bg='white')
        lal_Numb.grid(row=0,column=0,padx=5,pady=7,sticky=W)
        txt_Numb=ttk.Entry(upper_frame,textvariable=self.var_numb,width=22,font=('맑은 고딕',12,'bold'))
        txt_Numb.grid(row=0,column=1,padx=5,pady=7)

        # label : Name 성명
        lal_Name=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="성명:",bg='white')
        lal_Name.grid(row=0,column=2,padx=5,pady=7,sticky=W)
        txt_Name=ttk.Entry(upper_frame,textvariable=self.var_name,width=22,font=('맑은 고딕',12,'bold'))
        txt_Name.grid(row=0,column=3,padx=5,pady=7)

        # label : Gender 성별
        lal_Gender=Label(upper_frame,text='성별:',font=('맑은 고딕',12,'bold'),bg='white')
        lal_Gender.grid(row=0,column=4,padx=5,pady=7,sticky=W)
        combo_Gender=ttk.Combobox(upper_frame,textvariable=self.var_gender,font=('맑은 고딕',12,'bold'),width=17,state='readonly')
        combo_Gender['value']=('남','여', '기타')
        combo_Gender.current(0)
        combo_Gender.grid(row=0,column=5,padx=5,pady=7, sticky=W)

        # label : Regist 재적
        lal_Regist=Label(upper_frame,text='재적:',font=('맑은 고딕',12,'bold'),bg='white')
        lal_Regist.grid(row=0,column=6,padx=5,pady=7,sticky=W)
        combo_Regist=ttk.Combobox(upper_frame,textvariable=self.var_regist,font=('맑은 고딕',12,'bold'),width=17,state='readonly')
        combo_Regist['value']=('재학','전학', '자퇴', '퇴학')
        combo_Regist.current(0)
        combo_Regist.grid(row=0,column=7,padx=5,pady=7, sticky=W)

        # label Phone 전화번호
        lal_Phone=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="전화번호:",bg='white')
        lal_Phone.grid(row=1,column=0,padx=5,pady=7,sticky=W)
        txt_Phone=ttk.Entry(upper_frame,textvariable=self.var_phone,width=22,font=('맑은 고딕',12,'bold'))
        txt_Phone.grid(row=1,column=1,sticky=W,padx=2,pady=7)

        # label Middle 출신중학교
        lal_Middle=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="출신중학교:",bg='white')
        lal_Middle.grid(row=1,column=2,padx=5,pady=7,sticky=W)
        txt_Middle=ttk.Entry(upper_frame,textvariable=self.var_middle,width=22,font=('맑은 고딕',12,'bold'))
        txt_Middle.grid(row=1,column=3,sticky=W,padx=2,pady=7)

        # label Jumin1 주민번호1
        lal_Jumin1=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="주민번호1:",bg='white')
        lal_Jumin1.grid(row=1,column=4,padx=5,pady=7,sticky=W)
        txt_Jumin1=ttk.Entry(upper_frame,textvariable=self.var_jumin1,width=22,font=('맑은 고딕',12,'bold'))
        txt_Jumin1.grid(row=1,column=5,sticky=W,padx=2,pady=7)

        # label Jumin2 주민번호2
        lal_Jumin2=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="주민번호2:",bg='white')
        lal_Jumin2.grid(row=1,column=6,padx=5,pady=7,sticky=W)
        txt_Jumin2=ttk.Entry(upper_frame,textvariable=self.var_jumin2,width=22,font=('맑은 고딕',12,'bold'))
        txt_Jumin2.grid(row=1,column=7,sticky=W,padx=2,pady=7)

        # label Address 주소
        lal_Address=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="주소:",bg='white')
        lal_Address.grid(row=2,column=0,padx=5,pady=7,sticky=W)
        txt_Address=ttk.Entry(upper_frame,textvariable=self.var_address,width=22,font=('맑은 고딕',12,'bold'))
        txt_Address.grid(row=2,column=1,sticky=W,padx=2,pady=7)

        # label Post 우편번호
        lal_Post=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="우편번호:",bg='white')
        lal_Post.grid(row=2,column=2,padx=5,pady=7,sticky=W)
        txt_Post=ttk.Entry(upper_frame,textvariable=self.var_post,width=22,font=('맑은 고딕',12,'bold'))
        txt_Post.grid(row=2,column=3,sticky=W,padx=2,pady=7)

        # label Email 이메일
        lal_Email=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="이메일:",bg='white')
        lal_Email.grid(row=2,column=4,padx=5,pady=7,sticky=W)
        txt_Email=ttk.Entry(upper_frame,textvariable=self.var_email,width=22,font=('맑은 고딕',12,'bold'))
        txt_Email.grid(row=2,column=5,sticky=W,padx=2,pady=7)

        # label Fatname 부성명
        lal_Fatname=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="부성명:",bg='white')
        lal_Fatname.grid(row=2,column=6,padx=5,pady=7,sticky=W)
        txt_Fatname=ttk.Entry(upper_frame,textvariable=self.var_fatname,width=22,font=('맑은 고딕',12,'bold'))
        txt_Fatname.grid(row=2,column=7,sticky=W,padx=2,pady=7)

        # label Fatphone 부전화
        lal_Fatphone=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="부전화:",bg='white')
        lal_Fatphone.grid(row=3,column=0,padx=5,pady=7,sticky=W)
        txt_Fatphone=ttk.Entry(upper_frame,textvariable=self.var_fatphone,width=22,font=('맑은 고딕',12,'bold'))
        txt_Fatphone.grid(row=3,column=1,sticky=W,padx=2,pady=7)

        # label Motname 모성명
        lal_Motname=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="모성명:",bg='white')
        lal_Motname.grid(row=3,column=2,padx=5,pady=7,sticky=W)
        txt_Motname=ttk.Entry(upper_frame,textvariable=self.var_motname,width=22,font=('맑은 고딕',12,'bold'))
        txt_Motname.grid(row=3,column=3,sticky=W,padx=2,pady=7)

        # label Motphone 모전화
        lal_Motphone=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="모전화:",bg='white')
        lal_Motphone.grid(row=3,column=4,padx=5,pady=7,sticky=W)
        txt_Motphone=ttk.Entry(upper_frame,textvariable=self.var_motphone,width=22,font=('맑은 고딕',12,'bold'))
        txt_Motphone.grid(row=3,column=5,sticky=W,padx=2,pady=7)

        # label Club 특활부서
        lal_Club=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="특활부서:",bg='white')
        lal_Club.grid(row=3,column=6,padx=5,pady=7,sticky=W)
        txt_Club=ttk.Entry(upper_frame,textvariable=self.var_club,width=22,font=('맑은 고딕',12,'bold'))
        txt_Club.grid(row=3,column=7,sticky=W,padx=2,pady=7)

        # label Religion 종교
        lal_Religion=Label(upper_frame,text='종교:',font=('맑은 고딕',12,'bold'),bg='white')
        lal_Religion.grid(row=4,column=0,padx=5,pady=7,sticky=W)
        combo_Religion=ttk.Combobox(upper_frame,textvariable=self.var_religion,font=('맑은 고딕',12,'bold'),width=17,state='readonly')
        combo_Religion['value']=('기독교','불교', '천주교', '무교', '기타')
        combo_Religion.current(0)
        combo_Religion.grid(row=4,column=1,padx=5,pady=7, sticky=W)

        # label Hooby 취미
        lal_Hooby=Label(upper_frame,font=('맑은 고딕',12,'bold'),text="취미:",bg='white')
        lal_Hooby.grid(row=4,column=2,padx=5,pady=7,sticky=W)
        txt_Hooby=ttk.Entry(upper_frame,textvariable=self.var_hobby,width=22,font=('맑은 고딕',12,'bold'))
        txt_Hooby.grid(row=4,column=3,padx=5,pady=7,sticky=W)

        # 학생 사진 Frame
        self.photo_frame = LabelFrame(Main_frame,bd=0, relief=RIDGE, bg='white')
        self.photo_frame.place(x=1280, y=10, width=250, height=280)

        self.load_default_image()

        # Button Frame
        button_frame=Frame(upper_frame,bd=0,relief=RIDGE, bg='white')
        button_frame.place(x=0,y=234,width=1530,height=60)

        btn_add = Button(button_frame, command=self.add_data, text="저장", font=btn_font, width=14, bg=btn_hlb_bg, fg='white')
        btn_add.grid(row=0, column=0, padx=10)

        btn_update = Button(button_frame, command=self.update_data, text="수정", font=btn_font, width=14, bg=btn_hlb_bg, fg='white')
        btn_update.grid(row=0, column=1, padx=10)

        btn_delete = Button(button_frame, command=self.delete_data, text="삭제", font=btn_font, width=14, bg=btn_hlb_bg, fg='white')
        btn_delete.grid(row=0, column=2, padx=10)

        btn_clear = Button(button_frame, command=self.clear_data, text="리셋", font=btn_font, width=14, bg=btn_hlb_bg, fg='white')
        btn_clear.grid(row=0, column=3, padx=10)

        # down Frame
        down_frame = LabelFrame(Main_frame,bd=0, bg='red',relief=RIDGE, font=('맑은 고딕', 11, 'bold'), fg='black')
        down_frame.place(x=0, y=300, width=1530, height=470)

        # Search Frame
        search_frame = LabelFrame(down_frame,bd=0, relief=RIDGE, bg=rtf_bg, fg='white')
        search_frame.place(x=0, y=0, width=1530, height=60)

        # 검색 콤보박스 (학번, 성명, 전화번호로 검색)
        self.var_com_search = StringVar()
        com_txt_search = ttk.Combobox(search_frame, textvariable=self.var_com_search, state="readonly", font=('맑은 고딕', 14), width=10)
        com_txt_search['value'] = ('학번', '성명','전화번호')
        com_txt_search.current(0)
        com_txt_search.grid(row=0, column=1, padx=5, pady=10, sticky=W)

        # 검색 텍스트 입력
        self.var_search_txt = StringVar()
        txt_search = ttk.Entry(search_frame, textvariable=self.var_search_txt, width=20, font=btn_font)
        txt_search.grid(row=0, column=2, padx=5,pady=10)

        # 검색 버튼
        btn_search = Button(search_frame, text="검색",command=self.search_data, font=btn_font, width=14, bg='blue', fg='white')
        btn_search.grid(row=0, column=3, padx=5,pady=8)
        btn_ShowAll = Button(search_frame, text="전부출력",command=self.fetch_data, font=btn_font, width=14, bg='blue', fg='white')
        btn_ShowAll.grid(row=0, column=4, padx=5,pady=8)

        # # Table Frame
        table_frame=Frame(down_frame,bd=0,relief=RIDGE, bg="white")
        table_frame.place(x=0,y=60,width=1530,height=360)

        scroll_x=ttk.Scrollbar(table_frame,orient=HORIZONTAL)
        scroll_y=ttk.Scrollbar(table_frame,orient=VERTICAL)

        self.student_table=ttk.Treeview(table_frame,columns=('numb','name','gender','regist','phone','middle','jumin1','jumin2','address','post','email','fatname',
            'fatphone','motname','motphone','club','religion','hobby'),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)

        scroll_x.pack(side=BOTTOM,fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)

        scroll_x.config(command=self.student_table.xview)
        scroll_y.config(command=self.student_table.yview)

        self.student_table.heading('numb',text='학번')
        self.student_table.heading('name',text='성명')
        self.student_table.heading('gender',text='성별')
        self.student_table.heading('regist',text='재적')
        self.student_table.heading('phone',text='전화번호')
        self.student_table.heading('middle',text='출신중학교')
        self.student_table.heading('jumin1',text='주민1')
        self.student_table.heading('jumin2',text='주민2')
        self.student_table.heading('address',text='주소')
        self.student_table.heading('post',text='우편번호')
        self.student_table.heading('email',text='이메일')
        self.student_table.heading('fatname',text='부성명')
        self.student_table.heading('fatphone',text='부전화')
        self.student_table.heading('motname',text='모성명')
        self.student_table.heading('motphone',text='모전화')
        self.student_table.heading('club',text='특활부서')
        self.student_table.heading('religion',text='종교')
        self.student_table.heading('hobby',text='취미')

        self.student_table['show']='headings'

        self.student_table.column('numb',width=60)
        self.student_table.column('name',width=70)
        self.student_table.column('gender',width=50)
        self.student_table.column('regist',width=50)
        self.student_table.column('phone',width=100)
        self.student_table.column('middle',width=100)
        self.student_table.column('jumin1',width=60)
        self.student_table.column('jumin2',width=60)
        self.student_table.column('address',width=150)
        self.student_table.column('post',width=60)
        self.student_table.column('email',width=100)
        self.student_table.column('fatname',width=60)
        self.student_table.column('fatphone',width=100)
        self.student_table.column('motname',width=60)
        self.student_table.column('motphone',width=100)
        self.student_table.column('club',width=100)
        self.student_table.column('religion',width=80)
        self.student_table.column('hobby',width=70)

        self.student_table.pack(fill=BOTH,expand=1)
        # self.fetch_data()
        self.student_table.bind("<ButtonRelease-1>", self.get_cursor)
        self.add_data()

        # UnderBar Frame
        und_frame=Frame(self.root,bd=0,relief=RIDGE)
        und_frame.place(x=0,y=820,width=1530,height=60)

        lbl_school=Label(und_frame,font=('맑은 고딕',18,'bold'),text="유신고등학교 Since 1972   Tel : 070-5129-2979 | Fax : 031-211-1614",
            bg = lo_bg, fg = 'white')
        lbl_school.place(x=0, y=0, width=1530, height=60)

    def load_excel(self):
        if not os.path.exists(self.excel_file):
            self.wb = openpyxl.Workbook()
            self.ws = self.wb.active
            self.ws.append(["Numb", "Name", "Gender", "Regist", "Phone", "Middle", "Jumin1", "Jumin2", "Address","Post", "Email",
                            "Fatname", "Fatphone", "Motname", "Motphone", "Club", "Religion", "Hobby"])  # Add headers

            self.wb.save(self.excel_file)
        else:
            # Load the existing workbook
            self.wb = openpyxl.load_workbook(self.excel_file)
            self.ws = self.wb.active

        if not self.ws:
            raise Exception("워크시트가 초기화 되지 않음")

    def load_default_image(self):
        self.default_img_path = 'images/default_photo.jpg'
        img_4 = Image.open(self.default_img_path)
        img_4 = img_4.resize((230, 260), Image.LANCZOS)
        self.photo_img = ImageTk.PhotoImage(img_4)

        self.img_4 = Label(self.photo_frame, image=self.photo_img)
        self.img_4.place(x=5, y=10, width=230, height=260)

    def add_data(self):
        if not self.var_numb.get() or not self.var_name.get():
            if not self.is_first_input:
                messagebox.showerror("Error", "All fields are required")
            return
        self.is_first_input = False
        row = (self.var_numb.get(),self.var_name.get(),self.var_gender.get(),self.var_regist.get(),self.var_phone.get(),self.var_middle.get(),self.var_jumin1.get(),
            self.var_jumin2.get(),self.var_address.get(),self.var_post.get(),self.var_email.get(),self.var_fatname.get(),self.var_fatphone.get(),self.var_motname.get(),
            self.var_motphone.get(),self.var_club.get(),self.var_religion.get(),self.var_hobby.get())

        self.ws.append(row)
        self.wb.save(self.excel_file)
        self.fetch_data()
        self.clear_data()

    def update_data(self):
        try:
            selected_item = self.student_table.selection()

            if not selected_item:
                messagebox.showerror("Error", "Please select a record to update")
                return

            selected_item = selected_item[0]
            selected_values = self.student_table.item(selected_item, 'values')

            if not self.var_numb.get() or not self.var_name.get():
                messagebox.showerror("Error", "All fields (학번 and 성명) are required")
                return

            updated_row = (self.var_numb.get(),self.var_name.get(),self.var_gender.get(),self.var_regist.get(),self.var_phone.get(),self.var_middle.get(),self.var_jumin1.get(),
            self.var_jumin2.get(),self.var_address.get(),self.var_post.get(),self.var_email.get(),self.var_fatname.get(),self.var_fatphone.get(),self.var_motname.get(),
            self.var_motphone.get(),self.var_club.get(),self.var_religion.get(),self.var_hobby.get())

            for row_idx, excel_row in enumerate(self.ws.iter_rows(min_row=2, values_only=False)):
                if excel_row[0].value == selected_values[0]:  # Matching based on 학번
                    for col_idx, cell in enumerate(excel_row):
                        cell.value = updated_row[col_idx]  # Update cell value
                    break

            self.wb.save(self.excel_file)

            self.fetch_data()
            self.clear_data()

            messagebox.showinfo("Success", "Record updated successfully")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def load_student_image(self, student_number):
        image_path = f"images/{student_number}.jpg"

        if os.path.exists(image_path):
            img = Image.open(image_path)
        else:
            img = Image.open("images/default_photo.jpg")

        img = img.resize((230, 260), Image.LANCZOS)
        self.photo_img = ImageTk.PhotoImage(img)

        self.img_4.config(image=self.photo_img)
        self.img_4.image = self.photo_img

    def delete_data(self):
        selected_item = self.student_table.selection()[0]
        selected_values = self.student_table.item(selected_item, 'values')

        for row_idx, excel_row in enumerate(self.ws.iter_rows(min_row=2, values_only=False)):
            if excel_row[0].value == selected_values[0]:
                self.ws.delete_rows(row_idx + 2)
                break

        self.wb.save(self.excel_file)
        self.fetch_data()
        self.clear_data()

    def search_data(self):
        search_by = self.var_com_search.get()
        search_txt = self.var_search_txt.get().strip()

        if search_by == '학번':
            column_index = 0  #
        elif search_by == '성명':
            column_index = 1  #
        elif search_by == '전화번호':
            column_index = 4
        else:
            messagebox.showerror("Error", "Please select a valid field for searching!")
            return

        self.student_table.delete(*self.student_table.get_children())  # Clear the current table


        for row in self.ws.iter_rows(min_row=2, values_only=True):
            if search_txt.lower() in str(row[column_index]).lower():
                self.student_table.insert('', END, values=row)  #  END='end'

        if not self.student_table.get_children():
            messagebox.showinfo("Info", "No matching records found")

    def fetch_data(self):
        self.student_table.delete(*self.student_table.get_children())
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            self.student_table.insert('', END, values=row)

    def clear_data(self):
        self.var_numb.set("")
        self.var_name.set("")
        self.var_gender.set("")
        self.var_regist.set("")
        self.var_phone.set("")
        self.var_middle.set("")
        self.var_jumin1.set("")
        self.var_jumin2.set("")
        self.var_address.set("")
        self.var_post.set("")
        self.var_email.set("")
        self.var_fatname.set("")
        self.var_fatphone.set("")
        self.var_motname.set("")
        self.var_motphone.set("")
        self.var_club.set("")
        self.var_religion.set("")
        self.var_hobby.set("")

    def get_cursor(self, event):
        try:
            cursor_row = self.student_table.focus()
            contents = self.student_table.item(cursor_row)
            row = contents['values']

            if row:
                self.var_numb.set(row[0])
                self.var_name.set(row[1])
                self.var_gender.set(row[2])
                self.var_regist.set(row[3])
                self.var_phone.set(row[4])
                self.var_middle.set(row[5])
                self.var_jumin1.set(row[6])
                self.var_jumin2.set(row[7])
                self.var_address.set(row[8])
                self.var_post.set(row[9])
                self.var_email.set(row[10])
                self.var_fatname.set(row[11])
                self.var_fatphone.set(row[12])
                self.var_motname.set(row[13])
                self.var_motphone.set(row[14])
                self.var_club.set(row[15])
                self.var_religion.set(row[16])
                self.var_hobby.set(row[17])

                # Fetch the student number
                student_number = row[0]

                # Load the corresponding student image
                self.load_student_image(student_number)
            else:
                messagebox.showerror("Error", "선택된 Data에 자료 없음")
        except IndexError as e:
            messagebox.showerror("Error", f"데이터 접근 에러: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"오류 : {e}")


if __name__=="__main__":
    root = Tk()
    root.title('STUDENT MANAGEMENT')

    style = ttk.Style()
    style.configure("Treeview", font=("맑은 고딕", 10))

    app = student(root)
    # Run the program
    root.mainloop()
