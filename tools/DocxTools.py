import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, filedialog
from tkcalendar import Calendar
from datetime import date
import locale
from ReportEdit import ReportExporter
import os

class ReportDocxApp:
    def __init__(self):
        # Set up
        locale.setlocale(locale.LC_ALL, '')  # Set locale to system default

        # Get today's date
        self.today = date.today()
        # Format the date as "YYYY-MM-DD"
        self.default_date = self.today.strftime("%Y-%m-%d")
        
        # Create the window
        self.root = tk.Tk()
        self.root.title("Illumina Docx Report Maker")
        
        # 변수 지정
        self.SampleNamer_Window = None
        self.SampleID_List = []

        # Setting
        self.root.geometry("400x500+600+200")
        self.root.configure(bg='white')
        
        # protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.Input_Fields = {"ServceID":"",
                             "Client":"",
                             "Company":"",
                             "Platform":"",
                             "Type":"",
                             "Sample count":"",
                             "Library Kit":""}
        
        
        ## 필드 구성
        # Define the spacing between objects in the grid
        horizontal_spacing = 20
        vertical_spacing = 5

        # Input fields Service ID for the report
        self.Label_ServieID = tk.Label(self.root, text="Service ID", bg="white")
        self.Entry_ServieID = tk.Entry(self.root)
        self.Input_Fields["ServceID"] = self.Entry_ServieID
        
        self.Label_client = tk.Label(self.root, text="Client Name", bg="white")
        self.Entry_client = tk.Entry(self.root)
        
        self.Label_ClientCompany = tk.Label(self.root, text="Client Company", bg="white")
        self.Entry_ClientCompany = tk.Entry(self.root)
        
        ## Library Info
        # Platform of ComboBox widgets
        self.Var_Platform = ['Illumina Novaseq6000', 'Illumina NestSeq 500','Illumina MiSeq']
        self.Label_Platform = tk.Label(self.root, text="Platform", bg="white")
        self.Combobox_Platform = ttk.Combobox(self.root, values=self.Var_Platform)
        self.Combobox_Platform.set(self.Var_Platform[0])
        
        
        # ServiceType of ComboBox widgets
        self.Var_ServiceType = ['WGS', 'WTS/mRNA Seq', 'WES', 'Methyl', 'de novo', 'Seqeuncing Only', 'Single Cell RNA Sequencing']
        self.Label_ServiceType = tk.Label(self.root, text="ServiceType", bg="white")
        self.Combobox_ServiceType = ttk.Combobox(self.root, height=5, values=self.Var_ServiceType)
        self.Combobox_ServiceType.set(self.Var_ServiceType[0])

        # Labrary of ComboBox widgets
        self.Var_Labrary = self.Load_Library_Kit() # 라이브러리 불러오는 툴 만들기
        self.Label_Labrary = tk.Label(self.root, text="Labrary", bg="white")
        self.Combobox_Labrary = ttk.Combobox(self.root, height=15, values=self.Var_Labrary)
        self.Combobox_Labrary.set("Kit")

        # Library Format
        self.Var_Format = ["76-PE", "76-SE", "151-PE", "301-PE", "601-PE"]
        self.Label_Format = tk.Label(self.root, text="Labrary Format", bg="white")
        self.Combobox_Format = ttk.Combobox(self.root, height=5, values=self.Var_Format)
        self.Combobox_Format.set("151-PE")
        
        
        # Insert Size
        self.Var_Insert = ["550bp", "350bp", "155bp"]
        self.Label_Insert = tk.Label(self.root, text="Insert Size", bg="white")
        self.Combobox_Insert = ttk.Combobox(self.root, height=5, values=self.Var_Insert)
        self.Combobox_Insert.set("350bp")
        
        ## Sample Info
        # Count
        self.Entry_SampleCount =tk.IntVar()
        self.Label_SampleCount = tk.Label(self.root, text="Sample Count", bg="white")
        self.Entry_SampleCount = tk.Entry(self.root)
        self.Entry_SampleCount.insert(0, 5) # 임시

        # Sample Name 입력기
        self.SampleList = []
        self.Button_SampleName = tk.Button(self.root, text="Sample 입력", command=self.OpenSampleName)
        
        ## Minimum Sample Read or Base Pair
        self.TargetNumber_var = tk.StringVar()
        self.BaseVersion_var = tk.StringVar()
        
        self.BaseVersion_var.set("bp") # 목표생상량 기존 세팅 bp

        ## 생상량 UI
        self.Label_TargetNumver = tk.Label(self.root, text="목표생산량", bg="white")
        self.Entry_TargetNumber = tk.Entry(self.root, textvariable=self.TargetNumber_var)
        self.OptionMenu_Version = tk.OptionMenu(self.root, self.BaseVersion_var, "bp", "reads")
        
        ## Report Date - (Today, 입고일, Custom)
        self.Label_DeportDate = tk.Label(self.root, text="Report Date", bg="white")
       
        self.DateVar = tk.StringVar()
        self.DateVar.set(self.default_date)
        
        self.Entry_ReportDate = tk.Entry(self.root, textvariable=self.DateVar, validate="key", state=tk.DISABLED, width=12)
        self.Entry_ReportDate.configure(disabledforeground="black")  # Customize the text color in DISABLED state

        ## Report Date Calender 앱 실행
        self.Pick_Date_Button = tk.Button(self.root, text="날짜선택", command=self.pick_date)
        
        # CB variables
        self.CheckBox_DateVar1 = tk.IntVar()
        self.CheckBox_DateVar2 = tk.IntVar()
        self.CheckBox_DateVar3 = tk.IntVar()
        
        # CheckBox
        self.CheckBox_Date1 = tk.Checkbutton(self.root, text="Today", bg="white", variable=self.CheckBox_DateVar1, command=self.on_Clicked_DateCheckbox1) 
        self.CheckBox_Date2 = tk.Checkbutton(self.root, text="Service Date", bg="white", variable=self.CheckBox_DateVar2, command=self.on_Clicked_DateCheckbox2) 
        self.CheckBox_Date3 = tk.Checkbutton(self.root, text="Custom", bg="white", variable=self.CheckBox_DateVar3, command=self.on_Clicked_DateCheckbox3) 

        

        ### Grid 구성
        # Place the widgets in the window
        self.Label_ServieID.grid(row=0, column=0, pady=vertical_spacing, sticky=tk.E)
        self.Entry_ServieID.grid(row=0, column=1)
        
        self.Label_client.grid(row=1, column=0, pady=vertical_spacing, sticky=tk.E)
        self.Entry_client.grid(row=1, column=1)
        
        self.Label_ClientCompany.grid(row=2, column=0, pady=vertical_spacing, sticky=tk.E)
        self.Entry_ClientCompany.grid(row=2, column=1)
        
        self.Label_Platform.grid(row=3, column=0, sticky=tk.E)
        self.Combobox_Platform.grid(row=3, padx=horizontal_spacing, pady=vertical_spacing, column=1)
    
        self.Label_ServiceType.grid(row=4, column=0, sticky=tk.E)
        self.Combobox_ServiceType.grid(row=4, padx=horizontal_spacing, pady=vertical_spacing, column=1)

        self.Label_Labrary.grid(row=5, column=0, sticky=tk.E)
        self.Combobox_Labrary.grid(row=5, padx=horizontal_spacing, pady=vertical_spacing, column=1)

        self.Label_Format.grid(row=6, column=0, sticky=tk.E)
        self.Combobox_Format.grid(row=6, padx=horizontal_spacing, pady=vertical_spacing, column=1)

        self.Label_Insert.grid(row=7, column=0, sticky=tk.E)
        self.Combobox_Insert.grid(row=7, padx=horizontal_spacing, pady=vertical_spacing, column=1)
        
        self.Label_SampleCount.grid(row=8, column=0, sticky=tk.E)
        self.Entry_SampleCount.grid(row=8, padx=horizontal_spacing, pady=vertical_spacing, column=1)
        self.Button_SampleName.grid(row=8, column=2)

        self.Label_TargetNumver.grid(row=9, column=0, sticky=tk.E)
        self.Entry_TargetNumber.grid(row=9, column=1, padx=horizontal_spacing, pady=vertical_spacing)
        self.OptionMenu_Version.grid(row=9, column=2, sticky=tk.W)
        self.Entry_TargetNumber.bind("<KeyRelease>", self.Convert_Numver)

        self.Label_DeportDate.grid(row=10, column=0, sticky=tk.E)
        self.CheckBox_Date1.grid(row=10, column=1) ## 체크박스 생성
        self.CheckBox_Date2.grid(row=11, column=1)
        self.CheckBox_Date3.grid(row=12, column=1)
        self.Entry_ReportDate.grid(row=12, column=2) ## 날짜 입력기 캘린더
        self.CheckBox_DateVar1.set(1)
        self.Pick_Date_Button.grid(row=11, column=2) ## 날짜 입력기-자동변환
        self.Entry_ReportDate.bind("<KeyRelease>", self.format_date)


        # Create the button
        self.button = tk.Button(self.root, text="검수용 보고서 작성", command=self.create_report, sticky=None)
        self.button.grid(row=15, columnspan=2)
        
        self.create_Input_button()
        self.create_clear_button()
        

    ## 작동 기능 function        
    def on_closing(self):
        if messagebox.askokcancel("종료", "정말로 종료하시겠습니까?"):
            self.root.destroy()
            
            
    ## 필드 추가 기능 function
    def create_Input_button(self):
        InsertApp_button = tk.Button(self.root, text="적용", command=self.insert_input_serviceID)
        InsertApp_button.grid(row=0,column=2, sticky=tk.E, padx=6)    
        
    def create_clear_button(self):
        clear_button = tk.Button(self.root, text="초기화", command=self.clear_input_fields)
        clear_button.grid(row=0,column=2, sticky=tk.W)
            
    def insert_input_serviceID(self):
        None
    
    def clear_input_fields(self):
        self.Input_Fields["ServceID"].config(state=tk.NORMAL)
        self.Input_Fields["ServceID"].delete(0, tk.END)

    #### Sample 입력키 창 띄우기
    def OpenSampleName(self):
        if self.SampleNamer_Window:
            self.SampleNamer_Window.destroy()
        self.SampleNamer()
    
    def SampleNamer(self):
        if self.Entry_SampleCount.get() == "":
            messagebox.showwarning("빈 샘플 수", "샘플 수를 입력하세요")
            return None
            
        self.SampleNamer_Window = tk.Toplevel(self.root)
        self.SampleNamer_Window.title("샘플명 입력")
    
        # Setting
        self.SampleNamer_Window.geometry("+{}+{}".format(self.root.winfo_x() + self.root.winfo_width(), self.root.winfo_y()))
        
        self.num_fields_entry = None
        self.create_fields_button = None
        self.clear_fields_button = None
        self.entries = []

        label_Namer1 = tk.Label(self.SampleNamer_Window, text="샘플 이름을 입력하세요.\n기본값 Sample_$")
        label_Namer1.pack()

        num_fields = int(self.Entry_SampleCount.get())

        # Create new input fields
        for _ in range(num_fields):
            entry = tk.Entry(self.SampleNamer_Window, width=20)
            entry.pack()
            self.entries.append(entry)
            entry.bind("<Return>", lambda event, index=_: self.move_to_next_entry(event, index))
            
            if _ == 0:
                self.current_entry = entry
            
        self.SampleNamer_Window.bind("<Control-v>", self.handle_paste)
        
        self.Button_SampleSet = tk.Button(self.SampleNamer_Window, text="샘플명 확정", command=self.SampleListGet)
        self.Button_SampleSet.pack()
    
    
    def SampleListGet(self):
        self.SampleID_List = [entry.get() for entry in self.entries]
        self.SampleNamer_Window.destroy()
        
    def move_to_next_entry(self, event, index):
        if index < len(self.entries) - 1:
            self.current_entry = self.entries[index + 1]
            self.current_entry.focus_set()

    
    # 엑셀에서 복사해서 붙혀넣기 기능
    def handle_paste(self, event):
        content = self.root.clipboard_get()

        lines = content.split("\n")  # Split content by newline

        for i, line in enumerate(lines):
            if i < len(self.entries):
                self.entries[i].delete(0, tk.END)  # Clear existing entry
                self.entries[i].insert(0, line)  # Paste content into entry

    #### 라이브러리 DB
    def Load_Library_Kit(self):
        ngs_library_kits = [
                                "TruSeq DNA PCR-Free Library Prep Kit",
                                "NEBNext Ultra II DNA Library Prep Kit",
                                "KAPA HyperPrep Kit",
                                "Swift 2S Turbo DNA Library Kit",
                                "NEXTflex Rapid XP DNA-Seq Kit",
                                "SureSelect XT HS Library Prep Kit",
                                "Takara SMART-Seq Stranded Kit",
                                "Bioo Scientific NEXTflex ChIP-Seq Kit",
                                "Illumina Nextera XT DNA Library Prep Kit",
                                "Qiagen QIAseq FX DNA Library Kit",
                                "SMART-Seq v4 Ultra Low Input RNA Kit",
                                "10x Genomics Chromium Single Cell 3' Library Kit",
                                "Takara SMART-Seq v4 Pico Kit",
                                "Agilent SureSelect XT Human All Exon V7 Kit",
                                "IDT xGen Exome Research Panel v2"
                            ]
        return ngs_library_kits
    
    
    #### 체크 변경시 변경
    def Activate_DateEntry(self):
        
        if self.CheckBox_DateVar3.get() == 1:
            self.Entry_ReportDate.config(state=tk.NORMAL)
        else:
            self.Entry_ReportDate.config(state=tk.DISABLED)
            self.Entry_ReportDate.configure(disabledforeground="black")


    #### 생상랸 1000단위 표기기
    def Convert_Numver(self, *args):
        try:
            Data_out_commas = self.Entry_TargetNumber.get().replace(",", "")
            data = int(Data_out_commas)
        except (ValueError, TypeError):
            Data_Raw = self.Entry_TargetNumber.get()
            if Data_Raw:
                data = int(self.Entry_TargetNumber.get())
            else:
                return None
            
        self.Entry_TargetNumber.delete(0, tk.END)
        Output_Number = locale.format_string("%d", data, grouping=True)
        self.Entry_TargetNumber.insert(0, Output_Number)

            

    #### 날짜 자동 고정 입력기
    def format_date(self, *args):
        date = self.Entry_ReportDate.get()
        if len(date) == 4 and date.isdigit():
            self.Entry_ReportDate.delete(0, tk.END)
            self.Entry_ReportDate.insert(0, f"{date}-")
        elif len(date) == 8 and date[5] != '0' and date[5] != '1':
            self.Entry_ReportDate.delete(7, tk.END)
            self.Entry_ReportDate.insert(7, '0')
        elif len(date) == 8 and (date[6] == '3' or date[6] == '4' or date[6] == '5' or date[6] == '6' or date[6] == '7' or date[6] == '8' or date[6] == '9'):
            self.Entry_ReportDate.delete(7, tk.END)
            self.Entry_ReportDate.insert(7, '0')
        elif len(date) == 8 and (date[6] == '0' or date[6] == '1' or date[6] == '2'):
            self.Entry_ReportDate.insert(7, '-')
        elif len(date) > 10:
            self.Entry_ReportDate.delete(10, tk.END)
    
    #### CheckBox Set
    def on_Clicked_DateCheckbox1(self):
        if self.CheckBox_DateVar1.get() == 1:
            self.CheckBox_DateVar2.set(0)
            self.CheckBox_DateVar3.set(0)
            self.Activate_DateEntry()        

    def on_Clicked_DateCheckbox2(self):
        if self.CheckBox_DateVar2.get() == 1:
            self.CheckBox_DateVar1.set(0)
            self.CheckBox_DateVar3.set(0)
            self.Activate_DateEntry()        

    def on_Clicked_DateCheckbox3(self):       
        if self.CheckBox_DateVar3.get() == 1:
            self.CheckBox_DateVar1.set(0)
            self.CheckBox_DateVar2.set(0)
            self.Activate_DateEntry()           
    
    #### 날짜 선택 Calender APP
    def pick_date(self):
        # Create a new Tkinter window for the date picker
        self.Picker_Window = tk.Toplevel(self.root)
        self.Picker_Window.title("Report Date Picker")
        self.Picker_Window.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))

        # Create a Calendar widget and associate it with the date_var variable
        self.calendar = Calendar(self.Picker_Window, selectmode="day", datevar=self.default_date, firstweekday='sunday')
        self.calendar.pack()

        # Create a "Select" button to update the Entry widget with the selected date
        select_button = tk.Button(self.Picker_Window, text="Select", command=self.update_entry)
        select_button.pack()
    
    def update_entry(self):
        selected_date = self.calendar.selection_get().strftime("%Y-%m-%d")
        self.DateVar.set(selected_date)
        self.root.focus_set()
        self.Picker_Window.destroy()    
    
    ######
    ## Word Document Creating Functions
    def create_report(self):
        print("Creating Word Document")
        
        ## File path to the Word Document
        PWD = os.getcwd()
        ReportFilePath = filedialog.askopenfilename(defaultextension=".docx",
                                                    filetypes=[("Word Document", "*.docx")],
                                                    initialdir=PWD,
                                                    initialfile="TEST-Report.docx",
                                                    title="검수용 리포트 저장")
        
        Report = ReportExporter(ReportFilePath)
        Report.add_variable("Service ID", self.Entry_ServieID.get())
        
        ## Export Report
        Report.save_document()
        
        # ## Variable list
        # self.Entry_ServieID
        # self.Entry_client
        # self.Entry_ClientCompany
        # self.Combobox_Platform
        # self.Combobox_ServiceType
        # self.Combobox_Labrary
        # self.Combobox_Format
        # self.Combobox_Insert
        # self.Entry_SampleCount  
        # # Sample Name list
        # self.Entry_TargetNumber
        # self.OptionMenu_Version
        # self.CheckBox_Date1
        # self.CheckBox_Date2
        # self.CheckBox_Date3
        # self.CheckBox_DateVar1.set(1)
        # self.Entry_ReportDate.bind("<KeyRelease>", self.format_date)
    
    ###### 
    
    ## 실행 기능
    def run(self):
        self.root.mainloop()
        
if __name__ == "__main__":
    app = ReportDocxApp()
    app.run()
