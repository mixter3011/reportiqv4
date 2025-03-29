import tkinter as tk
from ui.ui import ReportIQ


def main():
    root = tk.Tk()
    app = ReportIQ(root)
    root.mainloop()

if __name__ == "__main__":
    main()
