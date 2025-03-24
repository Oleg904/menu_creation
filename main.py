import os
import shutil
from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showerror, showwarning, showinfo



home_dir = os.path.expanduser("~")

name_of_inst = ''

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
    root.geometry("450x400")

    # вывод текста в основном окне
    label = ttk.Label(root, text="Выберете файл типового меню", font=("Arial", 15))
    label.pack(padx=50, pady=60)

    # создание кнопки
    btn = ttk.Button(text="Выбрать", command=open_file)
    btn.pack(side=TOP, pady=30, ipadx=27, ipady=7)

    root.resizable(False, False)  # запрет изменения размеров окна

    root.protocol("WM_DELETE_WINDOW", finish)

    root.mainloop()

def menu_processing():
    try:  # проверка на наличие
        if len(os.listdir(f"{home_dir}/Desktop/Менюшки")) == 0:
            workbook = load_workbook(file_menu, read_only=True)
            sheet = workbook.active
            workbook2 = load_workbook("shablon.xlsx")
            sheet2 = workbook2.active
            sheet2[f"C{10}"] = "Пример"
            workbook2.save(f"{home_dir}/Desktop/Менюшки/Пример.xlsx")
            showinfo(title="Информация", message="Файлы меню созданы на Рабочем столе в папке Менюшки. Программу можно закрыть.")
            workbook.close()

        else:
            showinfo(title="Информация", message="В папке содержатся старые файлы меню. Эти файлы будут перезаписаны.")
            shutil.rmtree(f"{home_dir}/Desktop/Менюшки")
            os.mkdir(f"{home_dir}/Desktop/Менюшки")
            menu_processing()

    except BaseException as errors:
        if 'WinError 32' in str(errors):
            showinfo(title="Информация", message='Закройте файл меню и заново выберите файл типового меню.')
        else:
            showinfo(title="Информация", message=errors)


main_window()