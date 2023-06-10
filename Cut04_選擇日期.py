import tkinter as tk
from tkcalendar import Calendar


def 日期視窗位置(window):
    # 視窗長寬
    window_width = window.winfo_reqwidth()
    window_height = window.winfo_reqheight()

    # 螢幕長寬
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # 位置設定 = (螢幕長寬 - 視窗長寬) / 2
    x = int((screen_width - window_width) / 2)
    y = int((screen_height - window_height) / 2)

    # (視窗寬度 X 視窗高度 + 視窗X座標 + 視窗Y座標 )
    window.geometry(f"400x250+{x + 450}+{y}")


def 選擇日期():
    def tinker_日期選擇功能():
        # 在這裡聲明全域變數，就能讓外面一區的 選擇日期 函數
        global 目標日期
        selected_date = cal.selection_get()
        目標日期 = selected_date.strftime("%Y/%#m/%#d %H:%M")
        root.destroy()

    root = tk.Tk()

    root.title("選擇日期")

    cal = Calendar(root, selectmode="day")
    cal.pack()

    select_button = tk.Button(root, text="確定", command=tinker_日期選擇功能)
    select_button.pack()
    日期視窗位置(root)

    root.mainloop()
    return 目標日期
