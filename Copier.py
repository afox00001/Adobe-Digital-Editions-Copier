import re
import pyautogui
import win32gui
from pywinauto.findwindows    import find_window
from win32gui import SetFocus
import win32gui as wgui
import win32process as wproc
import win32api as wapi
import win32gui
from time import sleep
import pyperclip
import os
from fpdf import FPDF
import img2pdf
from PIL import Image
from pytesseract import pytesseract


#Set Console Window Positon
hwnd = win32gui.GetForegroundWindow()
win32gui.MoveWindow(hwnd, 2831, -1080, 1000, 500, True)

#Functions That Are Used To Convert The List Of Pictures To A PDF
def get_files(path):
    for file in os.listdir(path):
        if os.path.isfile(os.path.join(path, file)):
            yield file


def SaveImagesAsPDF(directory, pdfFileName, numberOfPages):
    images = []
    for i in range(numberOfPages):
        imageLocation = Image.open("./Book Pages/Page"+str(i+1)+".png").filename
        print(imageLocation)
        images.append(imageLocation)

    pdf = img2pdf.convert(images)
    with open(directory + pdfFileName, "wb") as pdfFile:
  
        pdfFile.write(pdf)

#Functions That Are Used To Set The Active Window
def win_enum_callback(hwnd, results):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        results.append(hwnd)

def print_list():
    handles = []
    win32gui.EnumWindows(win_enum_callback, handles)
    print('\n'.join(['%d\t%s' % (h, win32gui.GetWindowText(h)) for h in handles]))

def cycle_foreground():
    handles = []
    win32gui.EnumWindows(win_enum_callback, handles)
    for handle in handles:
        print(handle, win32gui.GetWindowText(handle))
        win32gui.SetForegroundWindow(handle)    

def SetActiveWindow(*argv):
    if not argv:
        window_name = "iCloud Passwords"
    else:
        window_name = argv[0]

    handle = wgui.FindWindow(None, window_name)
    print("Window `{0:s}` handle: 0x{1:016X}".format(window_name, handle))
    if not handle:
        print("Invalid window handle")
        return
    remote_thread, _ = wproc.GetWindowThreadProcessId(handle)
    wproc.AttachThreadInput(wapi.GetCurrentThreadId(), remote_thread, True)
    prev_handle = wgui.SetFocus(handle)

def check(s):
    pat = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    if re.match(pat,s):
        return True
    else:
        return False


#--------------Main Program----------------

def CopyEbookToPDF(bookTitle, numberOfPages):
    SetActiveWindow("Adobe Digital Editions - " + bookTitle)

    bookDirectory = r"./" + bookTitle + "/"
    for i in range(200):
        i += 1
        print("Takeing Screenshot Of Page #" + str(i))
        pyautogui.screenshot(bookDirectory + "Page" + str(i) + ".png", region=(681,85, 543, 913))
        sleep(0.1)
        pyautogui.hotkey("pagedown")
        sleep(0.2)
    SaveImagesAsPDF(bookDirectory, bookTitle + ".pdf", numberOfPages)
