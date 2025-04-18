import os
import datetime
import shutil
from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showinfo


home_dir = os.path.expanduser("~")

num_week_day = 6    # начальная строка для определения недели и дня недели

nutrition_calendar = ''     # путь к файлу календаря питания

file_menu = ''      # путь к файлу типового меню

how_day_menu = 0    # сколькидневное меню

def main_window():
    # открытие файла календаря питания
    def open_file_calendar():
        global nutrition_calendar
        nutrition_calendar = filedialog.askopenfilename(filetypes=(("EXCEL", ".xlsx"),))
        workbook3 = load_workbook(nutrition_calendar, read_only=True)  # выбор календаря
        sheet3 = workbook3.active  # выбор активного листа в календаре
        how_much_is_the_daily_menu(sheet3)
        workbook3.close()
        if nutrition_calendar != '':
            showinfo(title="Информация", message="Теперь выберите типовое меню.")
        else:
            showinfo(title="Информация", message="Вы не выбрали календарь питания. Выберите его заново.")

    # открытие файла типового меню
    def open_file_menu():
        global file_menu
        file_menu = filedialog.askopenfilename(filetypes=(("EXCEL", ".xlsx"),))
        if nutrition_calendar == '':
            showinfo(title="Информация", message="Вы не выбрали календарь питания. Сначала выберите его, а затем заново файл типового меню.")
        else:
            menu_processing()

    def finish():  # закрытие окна
        root.destroy()  # ручное закрытие окна и всего приложения

    # создается главное окно
    root = Tk()
    root.title("Создание меню")
    root.geometry("500x450")

    # вывод текста в основном окне
    label = ttk.Label(root, text="Выберете календарь питания, \nа затем файл типового меню", font=("Arial", 16))
    label.place(relx= 0.5, rely= 0.2, anchor=CENTER)

    # создание кнопки выбора типового меню
    btn = ttk.Button(text="Выбрать меню", command=open_file_menu)
    btn.place(relx= 0.5, rely= 0.5,height=40, width=180, anchor=CENTER)

    # создание кнопки выбора календаря питания
    btn2 = ttk.Button(text="Выбрать календарь питания", command=open_file_calendar)
    btn2.place(relx= 0.5, rely= 0.35,height=40, width=180, anchor=CENTER)

    label2 = ttk.Label(root, text="Создание файлов может занять до одной минуты.\nПрограмма сообщит, когда закончит.\nПо завершении, меню создадуться на Рабочем столе в папке 'Менюшки'",foreground="#126b62", font=("Arial", 10))
    label2.place(relx= 0.5, rely= 0.7, anchor=CENTER)

    # подпись внизу главного окна
    label3 = ttk.Label(root, text="By: Макаров Олег Николаевич МКОУ СОШ №11 г.Тавда", font=("Arial", 8))
    label3.place(relx= 0.3, rely= 0.95, anchor=CENTER)

    root.resizable(False, False)  # запрет изменения размеров окна

    root.protocol("WM_DELETE_WINDOW", finish)

    root.iconbitmap('files/image.ico')

    root.mainloop()

def how_much_is_the_daily_menu(sheet):   # определение сколько дней в меню
    start_reding = 4  # строка начала считывания дней в календаре
    read_column = 2 # столбец начала считывания
    global how_day_menu
    while True:
        while True:
            if not sheet.cell(row=start_reding, column=read_column).value is None and sheet.cell(row=start_reding, column=read_column).value > how_day_menu:
                how_day_menu = sheet.cell(row=start_reding, column=read_column).value
            if not sheet.cell(row=start_reding, column=read_column).value is None and sheet.cell(row=start_reding, column=read_column).value < how_day_menu:
                break
            if read_column == 32:
                break
            read_column += 1
        if start_reding == 13:
            break
        start_reding += 1
        read_column = 2

def cycle(row_of_sheet, sheet, sheet2):
    row_day_menu = 4    # строка начала вставки в ежедневное меню
    while True:
        if str(sheet.cell(row=row_of_sheet, column=4).value) == "Итого" or str(sheet.cell(row=row_of_sheet, column=4).value) == "итого":     # пропуск строки Итого
            sheet2.cell(row=row_day_menu, column=2).value = sheet.cell(row=row_of_sheet, column=4).value
            sheet2.cell(row=row_day_menu, column=5).value = None  # выход грамм
            sheet2.cell(row=row_day_menu, column=6).value = None  # цена
            sheet2.cell(row=row_day_menu, column=7).value = None  # калорийность
            sheet2.cell(row=row_day_menu, column=8).value = None  # белки
            sheet2.cell(row=row_day_menu, column=9).value = None  # жиры
            sheet2.cell(row=row_day_menu, column=10).value = None  # углеводы
            row_of_sheet += 1
            row_day_menu += 1
            continue
        sheet2.cell(row=row_day_menu,column=1).value = sheet.cell(row=row_of_sheet,column=3).value     # прием пищи
        sheet2.cell(row=row_day_menu, column=2).value = sheet.cell(row=row_of_sheet, column=4).value     # раздел
        sheet2.cell(row=row_day_menu, column=3).value = sheet.cell(row=row_of_sheet, column=11).value     # номер рецепта
        sheet2.cell(row=row_day_menu, column=4).value = sheet.cell(row=row_of_sheet, column=5).value     # наименование блюда
        sheet2.cell(row=row_day_menu, column=5).value = sheet.cell(row=row_of_sheet, column=6).value  # выход грамм
        sheet2.cell(row=row_day_menu, column=6).value = sheet.cell(row=row_of_sheet, column=12).value  # цена
        sheet2.cell(row=row_day_menu, column=7).value = sheet.cell(row=row_of_sheet, column=10).value  # калорийность
        sheet2.cell(row=row_day_menu, column=8).value = sheet.cell(row=row_of_sheet, column=7).value  # белки
        sheet2.cell(row=row_day_menu, column=9).value = sheet.cell(row=row_of_sheet, column=8).value  # жиры
        sheet2.cell(row=row_day_menu, column=10).value = sheet.cell(row=row_of_sheet, column=9).value  # углеводы
        if str(sheet.cell(row=row_of_sheet, column=3).value) == "Итого за день:" or str(sheet.cell(row=row_of_sheet, column=3).value) == "итого за день:":
            global num_week_day
            num_week_day = row_of_sheet + 1     # задание начальной строки для нового дня
            sheet2.cell(row=row_day_menu, column=5).value = None  # выход грамм
            sheet2.cell(row=row_day_menu, column=6).value = None  # цена
            sheet2.cell(row=row_day_menu, column=7).value = None  # калорийность
            sheet2.cell(row=row_day_menu, column=8).value = None  # белки
            sheet2.cell(row=row_day_menu, column=9).value = None  # жиры
            sheet2.cell(row=row_day_menu, column=10).value = None  # углеводы
            break
        else:   # спуск на строку ниже в типовом и ежедневном меню
            row_day_menu += 1
            row_of_sheet += 1

def menu_processing():
    try:  # проверка на наличие
        start_date = []  # дата начала действия типового меню
        if not os.path.exists(f"{home_dir}/Desktop/Менюшки"):   # проверка наличия папки с ежедневными меню и создание в случае отсутствия
            os.mkdir(f"{home_dir}/Desktop/Менюшки")
        if len(os.listdir(f"{home_dir}/Desktop/Менюшки")) == 0:
            workbook = load_workbook(file_menu, read_only=True)     # выбор файла типового меню
            sheet = workbook.active     # выбор активного листа
            # наименование учреждения
            school_name = sheet.cell(row=1,column=3).value
            # составление даты начала
            start_date.append(sheet.cell(row=3,column=10).value)     # год
            start_date.append(sheet.cell(row=3,column=9).value)     # месяц
            start_date.append(sheet.cell(row=3,column=8).value)     # день
            date = datetime.date(*start_date)
            current_date = date     # текущая дата меню
            while True:
                week = sheet.cell(row=num_week_day,column=1).value
                day_of_week = sheet.cell(row=num_week_day,column=2).value
                if current_date.isoweekday() == 6 and day_of_week != 6:      # если день выпадает на субботу
                    current_date += datetime.timedelta(2)
                elif current_date.isoweekday() == 7 and day_of_week != 7:    # если день выпадает на воскресенье
                    current_date += datetime.timedelta(1)
                workbook2 = load_workbook("files/shablon.xlsx")   # открытие шаблона
                sheet2 = workbook2.active     # выбор активного листа
                sheet2.cell(row=1, column=2).value = school_name    # вставка наименования учреждения в ежедневное меню
                sheet2.cell(row=1, column=10).value = current_date.strftime("%d.%m.%Y") # вставка даты в ежедневное меню
                cycle(num_week_day, sheet, sheet2)
                workbook2.save(f"{home_dir}/Desktop/Менюшки/{current_date.strftime("%Y-%m-%d")}-sm.xlsx")  # сохранение файла ежедневного меню
                current_date += datetime.timedelta(1)   # прибавление одних суток к дате
                if week * day_of_week == 10:      # завершение цикла после десятого дня
                    break
            workbook.close()
            showinfo(title="Информация", message="Файлы меню созданы. При необходимости, скорректируйте даты на ежедневных меню. Программу можно закрыть.")
            global nutrition_calendar
            nutrition_calendar = ''
        elif file_menu == '':
            showinfo(title="Информация", message="Вы не выбрали файл типового меню, выберите его снова.")
        else:
            showinfo(title="Информация", message="В папке содержатся старые файлы меню. Эти файлы будут перезаписаны.")
            shutil.rmtree(f"{home_dir}/Desktop/Менюшки")
            os.mkdir(f"{home_dir}/Desktop/Менюшки")
            menu_processing()
    except BaseException as errors:
        if 'WinError 32' in str(errors):
            showinfo(title="Информация", message="Закройте файл меню и заново выберите файл типового меню.")
        elif sheet.cell(row=3,column=10).value is None or sheet.cell(row=3,column=9).value is None or sheet.cell(row=3,column=8).value is None:
            workbook.close()
            showinfo(title="Информация", message="Не заполнена дата в типовом меню, пожалуйста заполните.")
        elif ValueError:
            print(errors)
            showinfo(title="Информация", message="В типовом меню введена некорректная дата. Пожалуйста, скорректируйте дату.")
        else:
            showinfo(title="Информация", message=str(errors))

main_window()