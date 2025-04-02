import os
import datetime
import shutil
from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showerror, showwarning, showinfo


home_dir = os.path.expanduser("~")

start_date = []     # дата начала действия типового меню

if not os.path.exists(f"{home_dir}/Desktop/Менюшки"):
    os.mkdir(f"{home_dir}/Desktop/Менюшки")

def main_window():
    # открытие файла типового меню
    def open_file():
        global file_menu
        file_menu = filedialog.askopenfilename(filetypes=(("EXCEL", ".xlsx"),))
        menu_processing()

    def finish():  # закрытие окна
        root.destroy()  # ручное закрытие окна и всего приложения


    # создается главное окно
    root = Tk()
    root.title("Создание меню")
    root.geometry("500x450")


    # вывод текста в основном окне
    label = ttk.Label(root, text="Выберете файл типового меню", font=("Arial", 15))
    label.place(relx= 0.5, rely= 0.2, anchor=CENTER)


    # создание кнопки
    btn = ttk.Button(text="Выбрать меню", command=open_file)
    btn.place(relx= 0.5, rely= 0.5,height=40, width=180, anchor=CENTER)

    label2 = ttk.Label(root, text="По завершении, меню создадуться на Рабочем столе в папке 'Менюшки'",foreground="#126b62", font=("Arial", 10))
    label2.place(relx= 0.5, rely= 0.7, anchor=CENTER)


    # подпись внизу главного окна
    label3 = ttk.Label(root, text="By: Макаров Олег Николаевич МКОУ СОШ №11 г.Тавда", font=("Arial", 8))
    label3.place(relx= 0.3, rely= 0.95, anchor=CENTER)

    root.resizable(False, False)  # запрет изменения размеров окна

    root.protocol("WM_DELETE_WINDOW", finish)

    root.mainloop()

def menu_processing():
    try:  # проверка на наличие
        if len(os.listdir(f"{home_dir}/Desktop/Менюшки")) == 0:
            workbook = load_workbook(file_menu, read_only=True)     # выбор файла типового меню
            sheet = workbook.active     # выбор активного листа
            # наименование учреждения
            school_name = sheet.cell(row=1,column=3).value
            # составление даты начала
            start_date.append(sheet.cell(row=3,column=10).value)
            start_date.append(sheet.cell(row=3,column=9).value)
            start_date.append(sheet.cell(row=3,column=8).value)
            date = datetime.date(*start_date)
            current_date = date     # текущая дата меню
            menu_day = 0    # номер дня меню
            while True:
                if current_date.isoweekday() == 6:      # если день выпадает на субботу
                    current_date += datetime.timedelta(2)
                    continue
                elif current_date.isoweekday() == 7:    # если день выпадает на воскресенье
                    current_date += datetime.timedelta(1)
                    continue
                menu_day += 1
                workbook2 = load_workbook("shablon.xlsx")   # открытие шаблона
                sheet2 = workbook2.active     # выбор активного листа
                sheet2.cell(row=1, column=2).value = school_name    # вставка наименования учреждения в ежедневное меню
                sheet2.cell(row=1, column=10).value = current_date.strftime("%d.%m.%Y") # вставка даты в ежедневное меню
                workbook2.save(f"{home_dir}/Desktop/Менюшки/{current_date.strftime("%Y-%m-%d")}-sm.xlsx")  # сохранение файла ежедневного меню
                current_date += datetime.timedelta(1)   # прибавление одних суток к дате
                if menu_day == 10:      # завершение цикла после десятого дня
                    break
            workbook.close()
            showinfo(title="Информация", message="Файлы меню созданы. При необходимости, скорректируйте даты на ежедневных меню. Программу можно закрыть.")
        else:
            showinfo(title="Информация", message="В папке содержатся старые файлы меню. Эти файлы будут перезаписаны.")
            shutil.rmtree(f"{home_dir}/Desktop/Менюшки")
            os.mkdir(f"{home_dir}/Desktop/Менюшки")
            menu_processing()

    except BaseException as errors:
        if 'WinError 32' in str(errors):
            showinfo(title="Информация", message="Закройте файл меню и заново выберите файл типового меню.")
        if 'openpyxl does not support  file format' in str(errors):
            showinfo(title="Информация", message="Вы не выбрали файл типового меню, выберите его снова.")
        else:
            showinfo(title="Информация", message=errors)

def cycle(row, col):
    while True:
        pass

main_window()

