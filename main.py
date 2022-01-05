import datetime
from tkinter import *
from tkinter import ttk
from functools import partial
from tkinter import filedialog

from excel_api import ExcelApi


class Window:
    def __init__(self):
        self.width = 824
        self.height = 400

        self.api = None
        self.current_date = datetime.datetime.today().date()

        window = Tk()
        window.title('Расписание занятий')
        window.geometry(f"{self.width}x{self.height}")
        self.window = window

        # Заголовок
        schedule_lbl = Label(window, text='Расписание занятий', font=('Arial Bold', 16))
        schedule_lbl.place(width=self.width, height=30)
        self.schedule_lbl = schedule_lbl

        date_string = self.current_date.strftime('%d.%m.%Y')
        today_lbl = Label(window, text=f"на {date_string}",  font=('Arial Bold', 14))
        today_lbl.place(width=self.width, height=30, y=30)
        self.today_lbl = today_lbl

        # Вчерашнее и завтрашнее числа
        yesterday_btn = Button(window, text='<', command=partial(self.increase_date, -1), bg='#555', fg='#ccc')
        yesterday_btn.place(width=20, height=30, x=314, y=30)
        self.tomorrow_btn = yesterday_btn

        tomorrow_btn = Button(window, text='>', command=partial(self.increase_date, 1), bg='#555', fg='#ccc')
        tomorrow_btn.place(width=20, height=30, x=490, y=30)
        self.tomorrow_btn = tomorrow_btn

        # Выбранный документ
        self.excel_file_path = ''

        excel_file_btn = Button(window, text='Выбрать документ', command=self.select_document, bg='#555', fg='#ccc')
        excel_file_btn.place(width=150, height=30, x=10, y=60)
        self.excel_file_btn = excel_file_btn

        excel_file_lbl = Label(window, text='Документ не выбран', font=('Arial Bold', 16), fg='red')
        excel_file_lbl.place(width=200, height=30, x=180, y=60)
        self.excel_file_lbl = excel_file_lbl

        # Поиск по преподавателю
        lecturer_lbl = Label(window, text='Преподаватель',  font=('Arial Bold', 14))
        lecturer_lbl.place(width=150, height=30, x=10, y=100)
        self.lecturer_lbl = lecturer_lbl

        self.lecturer_text = StringVar()
        lecturer_entry = Entry(window, width=10, textvariable=self.lecturer_text, state=DISABLED)
        lecturer_entry.place(width=150, height=30, x=180, y=100)
        lecturer_entry.focus()
        self.lecturer_entry = lecturer_entry

        find_by_lecturer_btn = Button(window, text='Найти', command=self.find_by_lecturer, bg='#555', fg='#ccc',
                                      state=DISABLED)
        find_by_lecturer_btn.place(width=150, height=30, x=350, y=100)
        self.find_by_lecturer_btn = find_by_lecturer_btn

        # Список дисциплин
        disciplines = ttk.Treeview(window)
        disciplines.place(x=10, y=150)

        disciplines['columns'] = ('Номер пары', 'Дисциплина', 'Учебная группа', 'Аудитория')
        disciplines.column('#0', width=0, stretch=NO)
        disciplines.column('#1', anchor=CENTER)
        disciplines.column('#2', anchor=CENTER)
        disciplines.column('#3', anchor=CENTER)
        disciplines.column('#4', anchor=CENTER)

        disciplines.heading('#0', text='', anchor=CENTER)
        disciplines.heading('#1', text='Номер пары', anchor=CENTER)
        disciplines.heading('#2', text='Дисциплина', anchor=CENTER)
        disciplines.heading('#3', text='Учебная группа', anchor=CENTER)
        disciplines.heading('#4', text='Аудитория', anchor=CENTER)
        self.disciplines = disciplines
        self.disciplines_rows = []

        window.mainloop()

    def increase_date(self, days):
        self.current_date += datetime.timedelta(days=days)
        date_string = self.current_date.strftime('%d.%m.%Y')
        self.today_lbl.configure(text=f"на {date_string}")

        # Обновляем список дисциплин
        self.find_by_lecturer()

    def select_document(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx;*.xlsm'),))

        if self.excel_file_path == '':
            self.excel_file_lbl.configure(text='Документ не выбран', font=("Arial Bold", 16), fg='red', justify='left')
            self.find_by_lecturer_btn.configure(state=DISABLED)
            self.lecturer_entry.configure(state=DISABLED)
        else:
            self.find_by_lecturer_btn.configure(state=NORMAL)
            self.lecturer_entry.configure(state=NORMAL)
            self.api = ExcelApi(self.excel_file_path)

            filename = self.excel_file_path.split('/')[-1]
            self.excel_file_lbl.configure(text=filename, font=("Arial Bold", 14), fg='green', justify='left')

    def find_by_lecturer(self):
        # remove added disciplines
        for i in range(len(self.disciplines_rows)):
            self.disciplines.delete(self.disciplines_rows[i])
        self.disciplines_rows.clear()

        # get disciplines of lecturer
        disciplines = self.api.get_discipline(self.current_date, self.lecturer_text.get())
        disciplines.sort(key=lambda x: x.number)

        for d in disciplines:
            row = self.disciplines.insert(parent='', index='end', values=(d.number, d.name, d.study_group, d.room))
            self.disciplines_rows.append(row)


if __name__ == '__main__':
    Window()
