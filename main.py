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
    def open_file():
        filedialog.askopenfile()
    # def secondary_window():
    #     window = Tk()
    #     window.title("Информация")
    #     window.geometry("450x80")
    #     label = ttk.Label(text="Просто", font=("Arial", 13))
    #     label.pack()

    def finish():
        root.destroy()  # ручное закрытие окна и всего приложения

    # создается главное окно
    root = Tk()
    root.title("Создание меню")
    root.geometry("450x400")

    btn = ttk.Button(text="Выберете файл типового меню", command=open_file)
    btn.pack(padx=100, pady=150, ipadx=10, ipady=10)

    root.resizable(False, False)  # запрет изменения размеров окна

    root.protocol("WM_DELETE_WINDOW", finish)

    root.mainloop()

main_window()

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

