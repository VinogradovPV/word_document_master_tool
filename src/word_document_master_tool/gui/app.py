import tkinter as tk

from .main_window import MainWindow


def main():
    root = tk.Tk()
    root.title("Word Document Master Tool")

    # Устанавливаем минимальный размер окна
    root.minsize(1000, 700)

    app = MainWindow(root)
    app.pack(fill=tk.BOTH, expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
