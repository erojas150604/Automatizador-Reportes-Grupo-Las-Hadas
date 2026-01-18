import tkinter as tk
from ui import AppUI
import sys

sys.setrecursionlimit(20000)

def main():
    root = tk.Tk()
    app = AppUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()