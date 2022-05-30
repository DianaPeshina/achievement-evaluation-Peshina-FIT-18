import os
import datetime
import math
import pandas as pd
import tkinter as tk
from tkinter import StringVar, IntVar, ttk
from tkinter import filedialog as fd

class GoalSetter(tk.Toplevel):
    def __init__(self, parent, title_set):
        super().__init__(parent)
        self.destroy_ind = 0
        self.title('Задание цели')
        self.label_num=len(title_set)
        self.minsize(200,self.label_num*27)
        label_list=[]
        self.text_list=[]
        textbox_list=[]
        for i in range(self.label_num):
            label_list.append(tk.Label(self, text=str(title_set[i])).pack())
            self.text_list.append(tk.StringVar(self))
            textbox_list.append(ttk.Entry(self,textvariable=self.text_list[i]).pack())
        btn = ttk.Button(self, text="Ок", command=self.values_check)
        btn.pack()
        self.protocol("WM_DELETE_WINDOW", self.handler)

    def handler(self):
        self.destroy_ind = 1
        self.destroy()

    def values_check(self):
        self.destroy_ind = 0
        for i in range(self.label_num):
            if not (('.' in self.text_list[i].get() and self.text_list[i].get().replace('.', '').isdigit())
                    or self.text_list[i].get().isdigit() 
                    or self.text_list[i].get() in ['зачет','незачет']):
                tk.messagebox.showinfo(title='Ошибка', message='Проверьте введенные данные')
                return
        self.destroy()
            
    def open(self):
        self.grab_set()
        self.wait_window()
        goal=[]
        for i in range(self.label_num):
            if '.' in self.text_list[i].get() and self.text_list[i].get().replace('.', '').isdigit():
                goal.append(float(self.text_list[i].get()))
            elif self.text_list[i].get().isdigit():
                goal.append(int(self.text_list[i].get()))
            elif self.text_list[i].get() == '': goal.append(1)
            else: goal.append(0)
        return goal, self.destroy_ind

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.data = pd.DataFrame()
        self.title('Приложение: Рейтинг успеваемости')
        self.geometry('900x600')
        self.minsize(900,600)
        self.frame_filedialog = tk.Frame(self,width=150,height=30)
        self.frame_view = tk.Frame(self, width=150, height=150)
        self.frame_buttons = tk.Frame(self, width=150, height=30)
        self.frame_filedialog.place(relx=0, rely=0, relwidth=1, relheight=0.08)
        self.frame_view.place(relx=0, rely=0.08, relwidth=1, relheight=0.84)
        self.frame_buttons.place(relx=0, rely=0.92, relwidth=1, relheight=0.08)
        self.open_button = ttk.Button(self.frame_filedialog, text='Открыть файл', command=self.select_file)
        self.text = tk.StringVar(self.frame_filedialog)
        self.open_textbox = ttk.Entry(self.frame_filedialog,textvariable=self.text)
        self.open_textbox.place(relx=0.01, rely=0.2, relwidth=0.6, relheight=0.6)
        self.open_button.place(relx=0.62, rely=0.2, relwidth=0.15, relheight=0.6)
        # место для отображения таблицы с данными
        self.table = ttk.Treeview(self.frame_view, show='headings')
        self.scroll_pane_y =ttk.Scrollbar(self.frame_view, command=self.table.yview)
        self.scroll_pane_x =ttk.Scrollbar(self.frame_view, command=self.table.xview, orient='horizontal')
        self.table.configure(yscrollcommand=self.scroll_pane_y.set, xscrollcommand=self.scroll_pane_x.set)
        self.scroll_pane_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scroll_pane_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)
        #конпки сохранения и оценки 
        self.no_goal = IntVar()
        self.no_pass = IntVar()
        self.flunk_button = ttk.Checkbutton(self.frame_buttons, text="Учитывать неуспевающих студентов", variable=self.no_pass).place(relx=0.23, rely=0.2,relheight=0.6)
        self.goal_button = ttk.Checkbutton(self.frame_buttons, text="Ввести цель вручную", variable=self.no_goal).place(relx=0.03, rely=0.2,relheight=0.6)
        self.rate_button = ttk.Button(self.frame_buttons,text='Рейтинг',command = self.rate_up)
        self.save_button = ttk.Button(self.frame_buttons, state=tk.DISABLED, text='Сохранить как',command=self.save_file_as)
        self.rate_button.place(relx=0.83, rely=0.2,relwidth=0.15, relheight=0.6)
        self.save_button.place(relx=0.65, rely=0.2,relwidth=0.15, relheight=0.6)
    
    # проверка имени файла в строке на существование и тип
    def name_check(self, file):
        if os.path.exists(file):
            filename, file_extension = os.path.splitext(file)
            if file_extension in ['.xls', '.xlsx', '.xlsm', '.xlsb']: return False
            else: return True
        else: return True

    # вывод таблицы на форму
    def table_update(self):
        self.table.delete(*self.table.get_children())
        headers=self.data.columns.values.tolist()
        self.table['columns'] = headers
        for header in headers:
            self.table.heading(header, text=header, anchor='center')
            self.table.column(header,anchor='center')
        for row in self.data.values.tolist():
            self.table.insert('', tk.END, values=row)

    # открыть выбранный файл
    def select_file(self):
        if self.text.get() == '' or self.name_check(self.text.get()):
            filetypes = (
                ('Excel Files', '*.xls *.xlsx *.xlsm *.xlsb'),
                ('All Files', '*.*')
            )
            filename = fd.askopenfilename(
                title='Выберите файл',
                initialdir=r'C:\\',
                filetypes=filetypes)
            self.text.set(filename)
        else: filename = self.text.get()
        
        if filename != '':
            self.save_button.configure(state=tk.DISABLED)
            self.rate_button.configure(state=tk.ACTIVE)
            # загрузка содержимого в датафрейм
            self.data = pd.read_excel(filename,0)
            self.table_update()

    # сохранение файла
    def save_file_as(self):
        filename = fd.asksaveasfilename(
            initialfile ="Рейтинг"+str(datetime.date.today()),
            defaultextension=".xlsx",
            initialdir=r'C:\\')
        with pd.ExcelWriter(filename) as writer:
            self.data.to_excel(writer)

    def target_proximity(self, vector, goal):
        goal
        result1 = result2 = 0
        ab=0
        a2=0
        b2=0
        size=len(goal)
        for i in range(size):
            if vector[i] == 0 and self.no_pass.get() == 0:
                return [0, 0]
            else: 
                ab=ab+goal[i]*vector[i]
                a2=a2+math.pow(goal[i],2)
                b2=b2+math.pow(vector[i],2)
        result1=ab/(math.sqrt(a2)*math.sqrt(b2)) 
        result2=ab*100/a2
        return [result1, result2]

    def rate_up(self):
        # переводим зачет и незачет в числовые эквиваленты
        # временный датафрейм для удобства вычислений
        df=self.data.copy()
        credit_map = {'незачет': 0, 'зачет': 1}
        df=df.replace(credit_map)
        first_column=''
        
        # заполняем пропущенные значения
        df = df.replace(r'^\s+$', pd.np.nan, regex=True)
        df.fillna(0,inplace=True)

        # проверяем присутствие только числовых значений
        columns = df.columns.tolist()
        types = df.dtypes.tolist()
        try:
            int_pos = types.index('int64')
        except ValueError:
            int_pos=len(columns)  
        try:
            float_pos = types.index('float64')
        except ValueError:
            float_pos=len(columns) 
        if float_pos>-1 and int_pos<float_pos:
            try:
                types.index('object',int_pos)
                tk.messagebox.showinfo(title='Ошибка', message='Данные предоставлены с ошибками. Проверьте источник информации')
                return
            except ValueError:
                first_column=columns[int_pos]
        elif int_pos>float_pos:
            try:
                types.index('object',float_pos)
                tk.messagebox.showinfo(title='Ошибка', message='Данные предоставлены с ошибками. Проверьте источник информации')
                return
            except ValueError:
                first_column=columns[float_pos]
        else: 
            tk.messagebox.showinfo(title='Ошибка', message='Данные предоставлены с ошибками. Проверьте источник информации')
            return
        
        # проверяем на отрицательные значения и отмечаем неуспевающих студентов
        for name in columns[columns.index(first_column):]:
            if not df[df[name]<0].empty :
                tk.messagebox.showinfo(title='Ошибка', message='В таблице присутствуют отрицательные значения.')
                return
            if self.no_pass.get() == 0 : df.loc[(df[name]==2), name] = 0

        # отделяем цель 
        if self.no_goal.get()==0:
            goal = df.loc[0][first_column:].values.tolist()
            self.data=self.data[1:].copy()
            df=df[1:].copy()
        else:
            window= GoalSetter(self, columns[columns.index(first_column):])
            goal, answer = window.open()
            if answer > 0: return
        
        # применяем формулы к данным
        self.data[['Приближенность к цели','Успешность, %']] = df.apply(lambda x: pd.Series(self.target_proximity(x.loc[first_column:].values.tolist(),goal)), axis =  1)
        # сортировка по убыванию
        self.data = self.data.sort_values(['Успешность, %','Приближенность к цели'],ascending=[False, False]).reset_index(drop=True)
        self.save_button.configure(state=tk.ACTIVE)
        self.rate_button.configure(state=tk.DISABLED)
        self.table_update()

if __name__ == "__main__":
    app = App()
    app.mainloop()