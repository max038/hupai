import time
import threading
import pyautogui
import pytesseract
import ctypes, ctypes.wintypes, win32con, win32gui
from datetime import datetime
import re


class HotKey(threading.Thread):

    def __init__(self):
        super(HotKey, self).__init__()

    def run(self):
        user32 = ctypes.windll.user32
        if not user32.RegisterHotKey(None, 99, win32con.MOD_CONTROL, ord("1")):
            raise RuntimeError
        try:
            msg = ctypes.wintypes.MSG()
            while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    if msg.wParam == 99:
                        x, y = pyautogui.position()
                        print(" ({}.{})".format(x, y))
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            user32.UnregisterHotKey(None, 1)


def TopMostMe():
    hwnd = win32gui.GetForegroundWindow()
    #(left, top, right, bottom) = win32gui.GetWindowRect(hwnd)
    #win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, left, top, right-left, bottom-top, 0)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 600, 300, 0)


if __name__=='__main__':
    TopMostMe()
    key = HotKey()
    key.start()
    key.join()
