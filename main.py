#!/usr/bin/env python3

import tkinter as tk
import src.gui as gui

if __name__ == "__main__":
    root = tk.Tk()
    app = gui.POSApplication(root)
    root.mainloop()