import tkinter as tk
from app.ui import AttendanceApp

def main():
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
