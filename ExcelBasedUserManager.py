from tkinter import *
import tkinter.filedialog as tkFile
import tkinter.ttk as ttk
import tkinter.font as font
import pandas as pd

class ExcelFileManager:
    '''
    Excel File에서 파일을 읽어 오거나 Data를 Excel File로 저장하는 역할을 하는 Class
    
    Field : 
        - PATH                      # 파일경로
    Method :
        + getPATH():String          # 파일의 경로를 반환하는 함수
        + openFile():DataFrame      # 파일을 열어 데이터를 DataFrame에 저장하여 반환하는 함수
        + saveFile(DataFrame):None  # DataFrame에 담긴 데이터를 바탕으로 Excel File로 저장
    '''  
    def __init__(self) -> None:
        self.__PATH = None
    
    # 파일의 경로를 반환하는 함수
    def getPATH(self) -> str:
        return self.__PATH

    # 파일을 열어 데이터를 DataFrame에 저장하여 반환하는 함수
    def openFile(self) -> pd.DataFrame:
        # Excel File 열기
        self.__PATH = tkFile.askopenfilename(initialdir="./", filetypes=(("*.xlsx","*xlsx"),("*.*","*")))
        if self.__PATH == "":
            return None 
        return  pd.read_excel(self.__PATH)

    # DataFrame에 담긴 데이터를 바탕으로 Excel File로 저장
    def saveFile(self, user_list:pd.DataFrame) -> None:
        try:
            user_list.to_excel(self.__PATH, index=False)
            return 1
        except:
            return 0
        


class UserListManager:
    '''
    User List를 관리하는 Class, User를 삭제하거나 추가하고, 또는 Property를 추가하는 기능을 수행한다.
    
    Filed:
        - user_list:DataFrame                               # User들의 데이터가 DataFrame형태로 저장될 변수
        - properties:list[String]                           # User들의 Property들을 정장하는 list
    Method()
        + setUserLsit(DataFrame):None                       # user_list의 값 설정
        + getUserList():DataFrame                           # user_list의 값 반환
        + getProperties():list[String]                      # properties값 반환
        + addProperty(String p_name, String default):None   # user list에 매개 변수로 받은 p_name값을 property의 이름으로 하고, default값을 값으로 가진 property추가    
        + addUser(dict):None                                # User에 대한 데이터를 받아 list에추가
        + deleteUser(int)                                   # index값을 받아 해당 index의 User 삭제
    '''
    def __init__(self) -> None:
        self.__user_list:pd.DataFrame = None
        self.__properties:list[str] = None

    def getUserList(self) -> pd.DataFrame:
        return self.__user_list
    
    def setUserList(self, data:pd.DataFrame) -> None:
        self.__user_list = data
        self.__properties = list(self.__user_list.columns)

    def getProperties(self):
        return self.__properties
    
    def deleteUser(self, index):
        self.__user_list = self.__user_list.drop(index)
        self.__user_list = self.__user_list.reset_index(drop=True)
        

    def addUser(self, data:dict):
        self.__user_list.loc[len(self.__user_list)] = data.values()


    def editUser(self, index, data:dict):
        selected_user = self.__user_list.iloc[index].copy()
        for key in data.keys():
            selected_user[key] = data[key]
        self.__user_list.iloc[index] = selected_user
            


class FilteringManager():
    '''
    User list의 필터링을 관리하는 Class
    Field:
    - selected_properties:list[String]                  # 필터링 옵션으로 선택된 Property들
    Method:
    + setSelectedProperties(list[String]):None          # selected_properties를 설정
    + getSelectedProperties():list[String]              # selected_properties의 값 반환
    + getValsOfProperty(DataFrame, String):list[String] # property에 대해 가질 수 있는 값의 종류, 단 가질 수 있는 값이 특정 수 이상인 경우 빈 리스트 반환 :
                                                        # Ex) proerty가 성별인 경우 : (남자, 여자) -> O, proerty가 전화번호인 경우 : (10^8) -> X
    + filteringUserList(DataFrame, String):DataFrame    # User list에 대해 filtering을 수행하여 그 결과를 반환                                                 
    '''
    def __init__(self) -> None:
        self.__selected_properties:list[str]

    def setSelectedProperties(self, properties:list[str]) -> None:
        self.__selected_properties = properties

    def getSelectedProperties(self) -> pd.DataFrame:
        return self.__selected_properties
    
    def getValsOfProperty(self, user_list:pd.DataFrame, property:str):
        vals = list(set(list(user_list[property].values)))
        if vals.__len__() > 4:
            return []
        else:
            return (["All"] + vals)

    def filteringUserList(self, user_list:pd.DataFrame, selected_data:dict) -> pd.DataFrame:

        data = user_list[self.__selected_properties]

        for properties in self.__selected_properties:
            if properties == "":
                continue
            if selected_data[properties] == "" or selected_data[properties] == "All":
                continue
            data = data[data[properties].astype(dtype=str)==selected_data[properties]]
        return data



            


class SortingManager():
    '''
    User list의 정렬을 담당하는 Class
    Field:
    Method:
        + sortingUserList(DataFrame, Stirng):DataFrame    # User list에 대해 매개변수로 받은 property를 기준으로 정렬 수행 후 정렬된 맴버리스트 반환
    '''
    def __init__(self) -> None:
        pass

class GUIManager():
    '''
    GUI를 담당하는 Manager
    Field:
        - window:window                             # 윈도우 창
        - file_manager:ExcelFileManager             # ExcelFileManger instance
        - user_list_manager:UserListManger          # UserListManger instance
        - filtering_manager:FilteringManager        # FilteringManager instance
        - sorting_manager:SortingManager            # SortingManager instance
        - showm_user_list:DataFrame                 # 현재 화면에 출력 중인 user list

    Method:
        - drawFileFrame():None                          # 파일과 관련된 GUI요소를 생성
        - drawFiltering():None                          # Filtering과 관련된 GUI요소 생성
        - drawSoring():None                             # Sorting과 관련된 GUI요소 생성
        - drawUserTable():None                          # User list가 출력하는 테이블을 생성
        - drawBottomFrame():None                        # GUI의 하단부분에 존재한는 요소(회원 추가, 속성추가 버튼 등) 생성
        - drawUserInfoWindow(DataFrame):None            # 유저의 살세정보를 출력하는 창을 생성 
        - drawUserAddWindow(DataFrame):None             # 유저의 정보를 입력받는 참을 생성
        - setDefault():None                             # 파일을 읽어왔을 때 기본적인 setting이 이루어지는 함수
        - openFile():None                               # 파일 열기와 관련된 동작을 하는 함수
        - saveFile():Nond                               # 파일 저장을 수행하는 함수
        - filtering():None                              # filtering과 관련된 함수, showm_user_list에 결과를 저장
        - sorting():None                                # 정렬과 관련된 함수, showm_user_list에 결과를 저장
        - fillTable():None                              # showm_user_list에 저장된 값을 가지고 Meber list Table을 채우는 함수
        - showUser(int):None                            # index에 해당하는 User의 세부정보 출력
        - addUser(dict):None                            # User를 list에 추가하는 함수
        - deleteUser(int):None                          # index에 해당하는 User를 삭제하는 함수
        - editUser(int):None  - editUser(int):None                            # index에 해당하는 User의 정보를 수정하는 함수
    '''
    def __init__(self, window) -> None:
        self.__window:Tk = window
        self.__window.geometry("1000x750+100+100")
        self.__window.resizable(True, True)
        self.__window.configure(background="White")
        self.__shwon_list:pd.DataFrame = None

        # Manager들 생성
        self.__file_manager = ExcelFileManager()
        self.__user_list_manager = UserListManager()
        self.__filtering_manager = FilteringManager()

        # GUI요소 생성
        self.__drawFileFrame()
        self.__drawFiltering()
        self.__drawSorting()
        self.__drawUserTable()
        self.__drawBottomFrame()

    def __setDefault(self, data:pd.DataFrame):
        self.__user_list_manager.setUserList(data)
        default_properties:list[str] = self.__user_list_manager.getProperties()

        for i in range(4):
            if i >= default_properties.__len__():
                default_properties.append("")
                self.__properties_cbox[i]["state"] = "disabled"
                self.__properties_cbox[i]["values"] = []
                self.__properties_cbox[i].set("")
                continue
            self.__properties_cbox[i]["state"] = "enabled"
            self.__properties_cbox[i]["values"] = default_properties
            self.__properties_cbox[i].set(default_properties[i])
            self.__setVals(i)

        for cbox in self.__vals_cbox:
            cbox.set("")
        self.__filtering()
        self.__sort_cbox["state"] = "enabled"
        self.__sort_cbox["values"] = self.__filtering_manager.getSelectedProperties()
        self.__sort_cbox.set(self.__filtering_manager.getSelectedProperties()[0])

        


    
    def __drawFileFrame(self) -> None:
        file_frame:LabelFrame = LabelFrame(self.__window, text="File", padx=5, pady=5, background="White", width=10)
        file_frame.pack(side=TOP, fill=X, padx=5, pady=5)
        file_select_btn:Button = Button(file_frame, text="Select File...", padx=5, command=self.__openFile)
        file_select_btn.pack(side=LEFT, padx=5, pady=5)
        self.__file_path_label:Label = Label(file_frame, text="PATH : Not Selceted File...", padx=10, background="White", anchor=W)
        self.__file_path_label.pack(side=LEFT, padx=5, pady=5)
        file_save_btn = Button(file_frame, text="Save", padx=5, command=self.__saveFile, width=10)
        file_save_btn.pack(side=RIGHT, padx=5, pady=5)
    def __openFile(self) -> None:
        data:pd.DataFrame = self.__file_manager.openFile()
        if type(data) == type(None):
            return
        self.__file_path_label.configure(text=self.__file_manager.getPATH())
        self.__setDefault(data)
    def __saveFile(self) -> None:
        result = self.__file_manager.saveFile(self.__user_list_manager.getUserList())
        result_win = Tk()
        result_win.title("Result")
        result_win.geometry("+300+300")
        result_label = Label(result_win, text="", padx=30, pady=30, font=font.Font(size=20))
        result_label.pack(side=TOP)
        ok_btn = Button(result_win, text="OK", padx=5, command=lambda: result_win.destroy(), width=10)
        ok_btn.pack(side=BOTTOM, padx=5, pady=5)
        if result == 1:
            result_label.configure(text="Save Sucesss")
        else:
            result_label.configure(text="Failed to Save")



    def __drawFiltering(self) -> None:
        filtering_frame:LabelFrame = LabelFrame(self.__window, text="Filtering", padx=5, pady=5, background="White")
        filtering_frame.pack(side=TOP, fill=X, padx=5, pady=5)
        
        self.__properties_cbox:list[ttk.Combobox] = []
        self.__vals_cbox:list[ttk.Combobox] = []

        for i in range(4):
            self.__properties_cbox.append(ttk.Combobox(filtering_frame, values=[], width=10, state="disabled"))
            self.__properties_cbox[-1].bind("<<ComboboxSelected>>", lambda event, i=i:self.__setVals(i))
            self.__properties_cbox[-1].pack(side=LEFT, padx=5, pady=5)
            self.__vals_cbox.append(ttk.Combobox(filtering_frame, values=[], width=10, state="disabled"))
            self.__vals_cbox[-1].pack(side=LEFT, padx=5, pady=5)

        apply_btn = Button(filtering_frame, text="Apply", padx=10, command=self.__filtering)
        apply_btn.pack(side=RIGHT, padx=5, pady=5)
    def __setVals(self, i):
        vals = self.__filtering_manager.getValsOfProperty(self.__user_list_manager.getUserList(), self.__properties_cbox[i].get())
        if vals == []:
            self.__vals_cbox[i]["state"] = "disabled"
            self.__vals_cbox[i]["values"] = []
            self.__vals_cbox[i].set("")
            return
        self.__vals_cbox[i]["state"] = "enabled"
        self.__vals_cbox[i]["values"] = vals
        self.__vals_cbox[i].set("")
    def __filtering(self) -> None:
        selected_data:dict = {}

        for i in range(4):
            s = self.__properties_cbox[i].get()
            if s == "":
                continue
            selected_data[s] = self.__vals_cbox[i].get()
 
        self.__filtering_manager.setSelectedProperties(list(selected_data.keys()))

        self.__shwon_list = self.__filtering_manager.filteringUserList(self.__user_list_manager.getUserList(), selected_data)
        self.__sort_cbox["values"] = self.__filtering_manager.getSelectedProperties()
        self.__sort_cbox.set(self.__filtering_manager.getSelectedProperties()[0])
        self.__fillTable()


    def __drawSorting(self) -> None:
        sorting_frame:LabelFrame = LabelFrame(self.__window, text="Sorting", padx=5, pady=5, background="White")
        sorting_frame.pack(side=TOP, fill=X, padx=5, pady=5)
        self.__sort_cbox = ttk.Combobox(sorting_frame, values=[], width=10, state="disabled")
        self.__sort_cbox.pack(side=LEFT, padx=5, pady=5)

        self.__sort_option = BooleanVar()
        ascending = Radiobutton(sorting_frame, text="Ascending", value=True, variable=self.__sort_option, background="White")
        ascending.pack(side=LEFT, padx=5, pady=5)
        descending = Radiobutton(sorting_frame, text="Descending", value=False, variable=self.__sort_option, background="White")
        descending.pack(side=LEFT, padx=5, pady=5)

        apply_btn = Button(sorting_frame, text="Apply", padx=10, command=self.__sorting)
        apply_btn.pack(side=RIGHT, padx=5, pady=5)
    def __sorting(self) -> None:
        self.__shwon_list = self.__shwon_list.sort_values(self.__sort_cbox.get(), ascending=self.__sort_option.get())
        self.__fillTable()
        
    def __drawUserTable(self) -> None:
        user_table_frame = LabelFrame(self.__window, text="MemberList", padx=10, pady=10, background="White")
        user_table_frame.pack(side=TOP, fill=X, padx=5, pady=5)
        self.__user_table = ttk.Treeview(user_table_frame, height=20)
        self.__user_table.pack()
    def __fillTable(self) -> None:
        selected_properties = self.__filtering_manager.getSelectedProperties()
        self.__user_table = ttk.Treeview(self.__user_table, height=20)
        self.__user_table.pack()
        self.__user_table["show"] = "headings"
        self.__user_table["columns"] = ["index"] + selected_properties
        self.__user_table["displaycolumns"] = ["index"] + selected_properties
        self.__user_table.column("index", width=50, anchor="center")
        self.__user_table.heading("index", text="index")
        for property in selected_properties:
            self.__user_table.column(property, anchor=E)
            self.__user_table.heading(property, text=property)
        
        for i in range(self.__shwon_list.__len__()):
            self.__user_table.insert("", "end", text="", values=[self.__shwon_list.iloc[[i]].index[0]]+list(self.__shwon_list.iloc[i].values), iid=self.__shwon_list.iloc[[i]].index[0])
        self.__user_table.bind("<Double-1>", self.__showUser)

    def __showUser(self, event) -> None:
        selected_index = int(self.__user_table.selection()[0])
        self.__drawUserInfoWindow(selected_index)
    def __drawUserInfoWindow(self, index):
        selected_user = self.__user_list_manager.getUserList().iloc[index]
        info_win = Tk()
        info_win.title("User Info")
        info_win.geometry("+150+150")
        info_win.configure(background="White")
        info_fram = LabelFrame(info_win, text="User Info", padx=10, pady=10, background="White", width=500)
        info_fram.pack(side=TOP, fill=X, padx=5, pady=5)
        for property in self.__user_list_manager.getProperties():
            if property == "":
                continue
            pv_frame = Frame(info_fram, background="White")
            pv_frame.pack(side=TOP, fill=X, padx=5, pady=5)
            property_label = Label(pv_frame, text=property, padx=10, background="White", anchor=W, width=15)
            property_label.pack(side=LEFT, padx=5, pady=5)
            mid_point = Label(pv_frame, text=":", padx=10, background="White", width=5)
            mid_point.pack(side=LEFT, padx=5, pady=5)
            val_label = Label(pv_frame, text=selected_user[property], padx=10, background="White", anchor=E)
            val_label.pack(side=RIGHT, padx=5, pady=5, fill=X)

        btn_frame = Frame(info_win, padx=0, pady=5, background="White")
        btn_frame.pack(side=TOP, fill=X, padx=0, pady=5)
        close_btn = Button(btn_frame, text="Close", padx=10, command=lambda :info_win.destroy())
        close_btn.pack(side=RIGHT, padx=5, pady=5)
        delete_btn = Button(btn_frame, text="Delete", padx=10, command=lambda :self.__deleteUser(index, info_win))
        delete_btn.pack(side=RIGHT, padx=5, pady=5)
        edit_btn = Button(btn_frame, text="Edit", padx=10, command=lambda :self.__drawEditWindow(index, info_win))
        edit_btn.pack(side=RIGHT, padx=5, pady=5)
    def __deleteUser(self, index, win:Tk):
        win.destroy()
        self.__user_list_manager.deleteUser(index)
        self.__setDefault(self.__user_list_manager.getUserList())

    def __drawEditWindow(self, index, win:Tk):
        win.destroy()
        selected_user = self.__user_list_manager.getUserList().iloc[index]
        edit_win = Tk()
        edit_win.title("User Info")
        edit_win.geometry("+150+150")
        edit_win.configure(background="White")
        edit_fram = LabelFrame(edit_win, text="Edit User Info", padx=10, pady=10, background="White", width=500)
        edit_fram.pack(side=TOP, fill=X, padx=5, pady=5)
        val_entries = {}
        for property in self.__user_list_manager.getProperties():
            if property == "":
                continue
            pv_frame = Frame(edit_fram, background="White")
            pv_frame.pack(side=TOP, fill=X, padx=5, pady=5)
            property_label = Label(pv_frame, text=property, padx=10, background="White", anchor=W, width=15)
            property_label.pack(side=LEFT, padx=5, pady=5)
            mid_point = Label(pv_frame, text=":", padx=10, background="White", width=5)
            mid_point.pack(side=LEFT, padx=5, pady=5)
            val_entry = Entry(pv_frame)
            val_entry.insert(0, selected_user[property])
            val_entry.pack(side=RIGHT, padx=5, pady=5, fill=X)
            val_entries[property] = val_entry
        btn_frame = Frame(edit_win, padx=0, pady=5, background="White")
        btn_frame.pack(side=TOP, fill=X, padx=0, pady=5)
        cancel_btn = Button(btn_frame, text="Cancel", padx=10, command=lambda :edit_win.destroy())
        cancel_btn.pack(side=RIGHT, padx=5, pady=5)
        apply_btn = Button(btn_frame, text="Apply", padx=10, command=lambda :self.__editUser(index, val_entries, edit_win))
        apply_btn.pack(side=RIGHT, padx=5, pady=5)


    def __editUser(self, index, val_entries:dict[str, Entry], win:Tk):
        def is_valid_float(element: str) -> bool:
            try:
                float(element)
                return True
            except ValueError:
                return False
        def is_valid_int(element: str) -> bool:
            if element[0] == "-":
                if element.__len__() > 2:
                    return element[1:].isnumeric()
            else:
                return element.isnumeric()
            
        for key in val_entries.keys():
            val = val_entries[key].get()
            if is_valid_int(val):
                val = int(val)
            elif is_valid_float(val):
                val = float(val)
            val_entries[key] = val
        self.__user_list_manager.editUser(index, val_entries)
        win.destroy()
        self.__setDefault(self.__user_list_manager.getUserList())
    
    def __drawAddUser(self):
        add_win = Tk()
        add_win.title("Add User Info")
        add_win.geometry("+150+150")
        add_win.configure(background="White")
        add_fram = LabelFrame(add_win, text="User Info", padx=10, pady=10, background="White", width=500)
        add_fram.pack(side=TOP, fill=X, padx=5, pady=5)
        val_entries = {}
        for property in self.__user_list_manager.getProperties():
            if property == "":
                continue
            pv_frame = Frame(add_fram, background="White")
            pv_frame.pack(side=TOP, fill=X, padx=5, pady=5)
            property_label = Label(pv_frame, text=property, padx=10, background="White", anchor=W, width=15)
            property_label.pack(side=LEFT, padx=5, pady=5)
            mid_point = Label(pv_frame, text=":", padx=10, background="White", width=5)
            mid_point.pack(side=LEFT, padx=5, pady=5)
            val_entry = Entry(pv_frame)
            val_entry.pack(side=RIGHT, padx=5, pady=5, fill=X)
            val_entries[property] = val_entry
        btn_frame = Frame(add_win, padx=0, pady=5, background="White")
        btn_frame.pack(side=TOP, fill=X, padx=0, pady=5)
        cancel_btn = Button(btn_frame, text="Cancel", padx=10, command=lambda :add_win.destroy())
        cancel_btn.pack(side=RIGHT, padx=5, pady=5)
        apply_btn = Button(btn_frame, text="Apply", padx=10, command=lambda :self.__addUser(val_entries, add_win))
        apply_btn.pack(side=RIGHT, padx=5, pady=5)

    def __drawBottomFrame(self) -> None:
        bottom_frame= Frame(self.__window, padx=5, pady=5, background="White")
        bottom_frame.pack(side=TOP, fill=X)
        add_btn = Button(bottom_frame, text="Add", padx=10, command=self.__drawAddUser, width=30)
        add_btn.pack(side=RIGHT)


    def __addUser(self, val_entries:dict[str, Entry], win:Tk):
        def is_valid_float(element: str) -> bool:
            try:
                float(element)
                return True
            except ValueError:
                return False
        def is_valid_int(element: str) -> bool:
            if element[0] == "-":
                if element.__len__() > 2:
                    return element[1:].isnumeric()
            else:
                return element.isnumeric()
            
        for key in val_entries.keys():
            val = val_entries[key].get()
            if is_valid_int(val):
                val = int(val)
            elif is_valid_float(val):
                val = float(val)
            val_entries[key] = val
        self.__user_list_manager.addUser(val_entries)

        win.destroy()
        self.__setDefault(self.__user_list_manager.getUserList())





if __name__ == "__main__":
    window = Tk()
    window.title("EcelBasedUserManager")
    app = GUIManager(window)
    window.mainloop()
