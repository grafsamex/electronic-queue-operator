from tkinter import ttk
from tkinter import *
from tkinter import filedialog
import sqlite3
import pandas as pd


class Programm:
    db_name = '\\\\srv\\NEW\\script\\baza.db'

    def __init__(self, window):

        self.wind = window
        self.wind.title('Редактирование очереди')

        # создание элементов для ввода слов и значений
        frame = LabelFrame(self.wind, text = 'Новый пациент')
        frame.grid(row = 0, column = 0, columnspan = 3, pady = 20)
        Label(frame, text = 'Время: ').grid(row = 1, column = 0)
        self.time = Entry(frame)
        self.time.focus()
        self.time.grid(row = 1, column = 1)
        Label(frame, text = 'ФИО: ').grid(row = 2, column = 0)
        self.fio = Entry(frame)
        self.fio.grid(row = 2, column = 1)
        ttk.Button(frame, text = 'Сохранить', command = self.add_fio).grid(row = 3, columnspan = 2, sticky = W + E)
        self.message = Label(text = '', fg = 'green')
        self.message.grid(row = 3, column = 0, columnspan = 2, sticky = W + E)
        # таблица слов и значений
        self.tree = ttk.Treeview(height = 10, columns = 2)
        self.tree.grid(row = 4, column = 0, columnspan = 2)
        self.tree.heading('#0', text = 'Время', anchor = CENTER)
        self.tree.heading('#1', text = 'ФИО', anchor = CENTER)

        # кнопки редактирования записей
        ttk.Button(text = 'Удалить', command = self.delete_fio).grid(row = 5, column = 0, sticky = W + E)
        ttk.Button(text = 'Изменить', command = self.edit_pacient).grid(row = 5, column = 1, sticky = W + E)
        ttk.Button(text = 'Загрузить из Excel', command=self.excel_export).grid(row = 6, column = 0, sticky = W + E, columnspan=3)
        # заполнение таблицы
        self.get_pacients()

    # подключение и запрос к базе
    def run_query(self, query, parameters = ()):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    # заполнение таблицы словами и их значениями
    def get_pacients(self):
        records = self.tree.get_children()
        for element in records:
            self.tree.delete(element)
        query = 'SELECT * FROM patients ORDER BY time DESC'
        db_rows = self.run_query(query)
        for row in db_rows:
            self.tree.insert('', 0, text = row[0], values = row[1])

    # валидация ввода
    def validation(self):
        return len(self.time.get()) != 0 and len(self.fio.get()) != 0
    # добавление нового слова
    def add_fio(self):
        if self.validation():
            query = 'INSERT INTO patients VALUES(?, ?)'
            parameters =  (self.time.get(), self.fio.get())
            self.run_query(query, parameters)
            self.message['text'] = 'Пациент {} добавлен в очередь'.format(self.fio.get())
            self.time.delete(0, END)
            self.fio.delete(0, END)
        else:
            self.message['text'] = 'введите время и ФИО'
        self.get_pacients()
    # удаление слова 
    def delete_fio(self):
        self.message['text'] = ''
        try:
            self.tree.item(self.tree.selection())['text'][0]
            print(self.tree.item(self.tree.selection())['text'][1])
        except IndexError as e:
            self.message['text'] = 'Выберите пациента, которого нужно удалить'
            return
        self.message['text'] = ''
        fio = self.tree.item(self.tree.selection())['values'][0]
        print(self.tree.item(self.tree.selection())['values'][0])
        print(fio)
        query = 'DELETE FROM patients WHERE fio = ?'
        self.run_query(query, (fio, ))
        self.message['text'] = 'Пациент успешно удален'
        self.get_pacients()
    # рeдактирование слова и/или значения
    def edit_pacient(self):
        self.message['text'] = ''
        try:
            self.tree.item(self.tree.selection())['values'][0]
        except IndexError as e:
            self.message['text'] = 'Выберите пациента для изменения'
            return
        time = self.tree.item(self.tree.selection())['text']
        old_fio = self.tree.item(self.tree.selection())['values'][0]
        self.edit_wind = Toplevel()
        self.edit_wind.title = 'Изменить пациента'

        Label(self.edit_wind, text = 'Прежнее время:').grid(row = 0, column = 1)
        Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = time), state = 'readonly').grid(row = 0, column = 2)
        
        Label(self.edit_wind, text = 'Новое время:').grid(row = 1, column = 1)
        # предзаполнение поля
        new_time = Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = time))
        new_time.grid(row = 1, column = 2)


        Label(self.edit_wind, text = 'Прежнее ФИО:').grid(row = 2, column = 1)
        Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = old_fio), state = 'readonly').grid(row = 2, column = 2)
 
        Label(self.edit_wind, text = 'Новое ФИО:').grid(row = 3, column = 1)
        # предзаполнение поля
        new_fio= Entry(self.edit_wind, textvariable = StringVar(self.edit_wind, value = old_fio))
        new_fio.grid(row = 3, column = 2)

        Button(self.edit_wind, text = 'Изменить', command = lambda: self.edit_records(new_time.get(), time, new_fio.get(), old_fio)).grid(row = 4, column = 2, sticky = W)
        self.edit_wind.mainloop()
    # внесение изменений в базу
    def edit_records(self, new_time, time, new_fio, old_fio):
        query = 'UPDATE patients SET time = ?, fio = ? WHERE time = ? AND fio = ?'
        parameters = (new_time, new_fio, time, old_fio)
        self.run_query(query, parameters)
        self.edit_wind.destroy()
        self.message['text'] = 'Пациент успешно изменен'
        self.get_pacients()
    #экспорт графика из Excel
    def excel_export(self):
        self.message['text'] = 'Экспортирую файл'
        excel_doc = filedialog.askopenfilename()
        if excel_doc == '':
            return
        df = pd.read_excel(excel_doc)
        
        for i in range(len(df)):
            query = 'INSERT INTO patients VALUES(?, ?)'
            parameters =  (df['Время'][i], df['ФИО'][i])
            self.run_query(query, parameters)
        self.get_pacients()
        self.message['text'] = ''


window = Tk()
application = Programm(window)
window.mainloop()