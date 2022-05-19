import os
import datetime
import math
import pandas as pd
import tkinter as tk
from tkinter import StringVar, ttk
from tkinter import filedialog as fd

data = pd.DataFrame()

root = tk.Tk()
root.title('Приложение: Рейтинг успеваемости')
root.geometry('900x600')
root.minsize(900,600)

frame_filedialog = tk.Frame(root,width=150,height=30)
frame_view = tk.Frame(root, width=150, height=150)
frame_buttons = tk.Frame(root, width=150, height=30)
frame_filedialog.place(relx=0, rely=0, relwidth=1, relheight=0.08)
frame_view.place(relx=0, rely=0.08, relwidth=1, relheight=0.84)
frame_buttons.place(relx=0, rely=0.92, relwidth=1, relheight=0.08)

# проверка имени файла в строке на существование и тип
def name_check(file):
    if os.path.exists(file):
        filename, file_extension = os.path.splitext(file)
        if file_extension in ['.xls', '.xlsx', '.xlsm', '.xlsb']: return False
        else: return True
    else: return True
# вывод таблицы на форму
def table_update():
    table.delete(*table.get_children())
    headers=data.columns.values.tolist()
    table['columns'] = headers
    for header in headers:
        table.heading(header, text=header, anchor='center')
        table.column(header,anchor='center')
    for row in data.values.tolist():
        table.insert('', tk.END, values=row)
# открыть выбранный файл
def select_file():
    if text.get() == '' or name_check(text.get()):
        filetypes = (
            ('Excel Files', '*.xls *.xlsx *.xlsm *.xlsb'),
            ('All Files', '*.*')
        )
        filename = fd.askopenfilename(
            title='Выберите файл',
            initialdir=r'C:\\',
            filetypes=filetypes)
        text.set(filename)
    else: filename = text.get()
    
    if filename != '':
        global data
        save_button.configure(state=tk.DISABLED)
        rate_button.configure(state=tk.ACTIVE)
        # загрузка содержимого в датафрейм
        data = pd.read_excel(filename,0)
        table_update()
# сохранение файла
def save_file_as():
    filename = fd.asksaveasfilename(
        initialfile ="Рейтинг"+str(datetime.date.today()),
        defaultextension=".xlsx",
        initialdir=r'C:\\')
    with pd.ExcelWriter(filename) as writer:
        data.to_excel(writer)
    

def rulew(vector, goal):
    goal
    result = 0
    ab=0
    a2=0
    b2=0
    size=len(goal)
    for i in range(size):
        ab=ab+goal[i]*vector[i]
        a2=a2+math.pow(goal[i],2)
        b2=b2+math.pow(vector[i],2)
    result=ab/(math.sqrt(a2)*math.sqrt(b2)) 
    return result

def ruleb(vector, goal):
    goal
    result = 0
    ab=0
    a2=0
    size=len(goal)
    for i in range(size):
        ab=ab+goal[i]*vector[i]
        a2=a2+math.pow(goal[i],2)
    result=ab*100/a2 
    return result

def rate_up():
    global data

    credit_map = {'незачет': 2, 'зачет': 3}
    data=data.replace(credit_map)

    goal = data.loc[0]['экзамен1':].values.tolist()
    temp=newdata=data[1:].copy()
    newdata['w'] = temp.apply(lambda x: rulew(x.loc['экзамен1':].values.tolist(), goal), axis =  1)
    newdata['b'] = temp.apply(lambda x: ruleb(x.loc['экзамен1':].values.tolist(), goal), axis =  1)
    data = newdata.sort_values(['b','w'],ascending=[False, False])
    save_button.configure(state=tk.ACTIVE)
    rate_button.configure(state=tk.DISABLED)
    table_update()

# строка для ввода имени файла и кнопка для вызова файлового диалога
open_button = ttk.Button(frame_filedialog, text='Открыть файл', command=select_file)
text = tk.StringVar(frame_filedialog)
open_textbox = ttk.Entry(frame_filedialog,textvariable=text)
open_textbox.place(relx=0.01, rely=0.2, relwidth=0.6, relheight=0.6)
open_button.place(relx=0.62, rely=0.2, relwidth=0.15, relheight=0.6)

# место для отображения таблицы с данными
table = ttk.Treeview(frame_view, show='headings')
scroll_pane_y =ttk.Scrollbar(frame_view, command=table.yview)
scroll_pane_x =ttk.Scrollbar(frame_view, command=table.xview, orient='horizontal')
table.configure(yscrollcommand=scroll_pane_y.set, xscrollcommand=scroll_pane_x.set)
scroll_pane_y.pack(side=tk.RIGHT, fill=tk.Y)
scroll_pane_x.pack(side=tk.BOTTOM, fill=tk.X)
table.pack(expand=tk.YES, fill=tk.BOTH)

#конпки сохранения и оценки 
rate_button = ttk.Button(frame_buttons,text='Оценить',command = rate_up)
save_button = ttk.Button(frame_buttons, state=tk.DISABLED, text='Сохранить как',command=save_file_as)
rate_button.place(relx=0.83, rely=0.2,relwidth=0.15, relheight=0.6)
save_button.place(relx=0.65, rely=0.2,relwidth=0.15, relheight=0.6)

root.mainloop()