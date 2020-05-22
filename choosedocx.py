# -*- coding: utf-8 -*-
'''
* @author UUGTech
* @copyright 2020 by UUGTech
'''
import os,sys
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

def choose():
    root = Tk()
    root.withdraw()
    fType = [("","*.docx")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    messagebox.showinfo('French Fries Paradise','人質にするwordファイル(.docx)を選んでください')


    f = filedialog.askopenfilename(filetypes = fType,initialdir = iDir)
    return f