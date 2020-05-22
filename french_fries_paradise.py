# -*- coding: utf-8 -*-
'''
* @author UUGTech
* @copyright 2020 by UUGTech
'''
import win32com.client
import time
import docx
import tkinter
import os
import sys
import keyboard
import pyautogui
from playsound import playsound
import threading


# make the word file (.docx) a french fries paradise
def french_fries_paradise(f):
    # don't change these charactors
    through_list = [" ",
                    "　",
                    "/",
                    "(",
                    "（",
                    ")",
                    "）",
                    ",",
                    ".",
                    ":",
                    ";",
                    "\'",
                    "\"",
                    "\\",
                    "\r",
                    chr(int(0x0c))]
    # open Word
    Application = win32com.client.Dispatch("Word.Application")
    # make the app visible
    Application.visible = True
    # open the document
    doc = Application.Documents.Open(f)
    # charators -> french fries
    i = 0
    while i < doc.Content.End:
        if keyboard.is_pressed('esc'):
            sys.exit()
        time.sleep(0.01)
        if str(doc.Range(i,i+1).Text) in through_list:
            i+=1
            continue
        else:
            doc.Range(i,i+1).Font.Size /= 2
            doc.Range(i,i+1).Text = chr(int(0x1f35f))
            i+=2
    # close the document
    doc.Close(SaveChanges=-1)
    # quit
    Application.Quit()
    sys.exit()


# play sound
def sound_func():
    sound = os.path.dirname(os.path.abspath(__file__)) + "/SE.mp3"
    while(True):
        playsound(sound)
        time.sleep(1)


# return full path of this file
def echo_path():
    return os.path.abspath(__file__)


# ---------- main ---------- #
if __name__ == "__main__":
    pyautogui.press("down") # activate monitor
    args = sys.argv
    sound_thread = threading.Thread(target=sound_func)
    sound_thread.setDaemon(True)
    sound_thread.start()
    french_fries_paradise(args[1])