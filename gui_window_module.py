import os
from openpyxl import load_workbook
from tkinter import messagebox
import tkinter as tk
from tkinter import Tk, ttk, Text,filedialog
class Gui(tk.Tk):
    def __init__(self,height=250, width=300, pdx=10, pdy=10):
        super().__init__()
        w_hgt=height; w_wdt=width; txt_padx=pdx; txt_pady=pdy
        self.title('Read & Sort Component')
        # scr_wdt = self.winfo_screenwidth(); scr_hgt = self.winfo_screenheight()

        self.resizable(False, False)
        scr_wdt = self.winfo_screenwidth()
        scr_hgt = (self.winfo_screenheight())
        centr_x = int((scr_wdt - w_wdt)/2)
        centr_y = int((scr_hgt - w_hgt)/2)
        self.geometry(f'{w_wdt}x{w_hgt}+{centr_x}+{centr_y}')

    #Buttons
        self.bt_first=tk.Button(self, text='First', width=15, height=4)
        self.bt_second=tk.Button(self, text="Second ", width=15, height=4 )
        self.bt_third=tk.Button(self, text='Third', width=15, height=4)
        self.bt_four=tk.Button(self, text='Four', width=15, height=4)
        self.lbl_name=tk.Label(self, text='Caption', width=10, height=1)
        self.txt_text=tk.Text(self, width=5, height=1, borderwidth=1)
    #Grid

        self.bt_first.grid(column=0, row=0, padx=txt_padx, pady=txt_pady)
        self.bt_second.grid(column=1, row=0, padx=txt_padx, pady=txt_pady)
        self.bt_third.grid(column=0, row=1, padx=txt_padx, pady=txt_pady )
        self.bt_four.grid(column=1, row=1, padx=txt_padx, pady=txt_pady, stick='w')
        self.lbl_name.grid(column=0, row=2, padx=txt_padx, pady=txt_pady)
        self.txt_text.grid(column=0,row=3, padx=txt_padx, stick='w' )


