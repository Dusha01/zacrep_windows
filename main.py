import tkinter as tk
import pywintypes
import win32gui
import win32con
import win32process
import psutil

class Main_win():
    def __init__(self, win):
        self.master = win
        win.title('Xd')

        self.label = tk.Label(win, text='Выбери окно: ')
        self.label.pack()

        self.windows_listbox = tk.Listbox(win, width=50, height=10)
        self.windows_listbox.pack()

        self.refresh_win_bt = tk.Button(win, text='Обновить', command=self.populate_window_list)
        self.refresh_win_bt.pack()

        self.pin_button = tk.Button(win, text="Закрепить окно", command=self.pin_window)
        self.pin_button.pack()

        self.unpin_button = tk.Button(win, text="Открепить окно", command=self.unpin_window)
        self.unpin_button.pack()

        self.status_label = tk.Label(win, text="Ожидание выбора окна")
        self.status_label.pack()

        self.pinned_windows = {}

        self.populate_window_list()

    def populate_window_list(self):
        self.windows_listbox.delete(0, tk.END)

        def enum_windows_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
                    try:
                        process = psutil.Process(process_id)
                        process_name = process.name()
                    except psutil.NoSuchProcess:
                        process_name = "Неизвестный процесс"

                    self.windows_listbox.insert(tk.END, f"{title}  (PID: {process_id}, Process: {process_name}) - HWND: {hwnd}")

        win32gui.EnumWindows(enum_windows_callback, None)

    def pin_window(self):
        selection = self.windows_listbox.curselection()
        if not selection:
            self.status_label.config(text="Статус: Сначала выберите окно.")
            return

        selected_window = self.windows_listbox.get(selection[0])
        try:
            hwnd = int(selected_window.split("HWND: ")[1])
        except IndexError:
            self.status_label.config(text="Обновите список окон.")
            return

        if hwnd in self.pinned_windows:
             self.status_label.config(text="Это окно уже закреплено.")
             return

        try:
            original_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)

            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

            self.pinned_windows[hwnd] = original_style

            self.status_label.config(text=f"Окно HWND {hwnd} закреплено.")

        except pywintypes.error as e:
            self.status_label.config(text=f"Ошибка при закреплении окна: {e}")

    def unpin_window(self):
        selection = self.windows_listbox.curselection()
        if not selection:
            self.status_label.config(text="Сначала выберите окно.")
            return

        selected_window = self.windows_listbox.get(selection[0])
        try:
            hwnd = int(selected_window.split("HWND: ")[1])
        except IndexError:
            self.status_label.config(text="Ошибка разбора HWND. Обновите список окон.")
            return

        if hwnd in self.pinned_windows:
            try:
                original_style = self.pinned_windows[hwnd]
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, original_style)
                win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)

                del self.pinned_windows[hwnd]

                self.status_label.config(text=f"Окно HWND {hwnd} откреплено.")
            except pywintypes.error as e:
                self.status_label.config(text=f"Ошибка при откреплении окна: {e}")
        else:
            self.status_label.config(text="Это окно не закреплено.")

root = tk.Tk()
my_gui = Main_win(root)
root.mainloop()