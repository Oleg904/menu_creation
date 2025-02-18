import os
from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showerror, showwarning, showinfo



home_dir = os.path.expanduser("~")

if not os.path.exists(f"{home_dir}/Desktop/Менюшки"):
    os.mkdir(f"{home_dir}/Desktop/Менюшки")

wb2 = load_workbook("shablon.xlsx")
sheet2 = wb2.active

def main_window():
    # открытие файла типового меню
    def open_file():
        # file_menu = filedialog.askopenfile(filetypes=(("EXCEL", ".xlsx"),))
        global file_menu
        global  wb
        global sheet
        file_menu = filedialog.askopenfilename(filetypes=(("EXCEL", ".xlsx"),))
        wb = load_workbook(file_menu, read_only=True)
        sheet = wb.active
        showinfo(title="Информация", message=sheet["A6"].value)

    # def secondary_window():
    #     window = Tk()
    #     window.title("Информация")
    #     window.geometry("450x80")
    #     label = ttk.Label(text="Просто", font=("Arial", 13))
    #     label.pack()

    def finish():
        print(sheet["A6"].value)
        wb.close()
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

# print(os.path.exists(f"{home_dir}/Desktop/Меню"))
# print(os.path.isfile(f"{home_dir}/Desktop/Меню/Пример.xlsx"))
def menu_processing():
    try:
        if len(os.listdir(f"{home_dir}/Desktop/Менюшки")) == 0:
            sheet2[f"C{10}"] = "Пример"
            wb2.save(f"{home_dir}/Desktop/Менюшки/Пример.xlsx")
        else:
            showinfo(title="Информация", message="В папке содержатся старые файлы меню. Эти файлы будут перезаписаны.")

    except BaseException as e:
        print(e)
    else:
        showinfo(title="Информация", message="Файлы меню созданы на Рабочем столе в папке Менюшки")

main_window()