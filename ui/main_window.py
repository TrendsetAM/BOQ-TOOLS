"""
BOQ Tools Main Window
Comprehensive GUI for Excel file processing and BOQ analysis
"""

import tkinter as tk

class MainWindow:
    def __init__(self, processor):
        self.processor = processor
        self.root = tk.Tk()
        self.root.title("BOQ Tools")

    def run(self):
        self.root.mainloop() 