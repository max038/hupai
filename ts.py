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
        self.timesync = None

    def run(self):
        user32 = ctypes.windll.user32
        if not user32.RegisterHotKey(None, 99, win32con.MOD_CONTROL, ord("1")):
            raise RuntimeError
        try:
            pos = None
            msg = ctypes.wintypes.MSG()
            while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    if msg.wParam == 99:
                        x, y = pyautogui.position()
                        if not pos:
                            pos = (x, y)
                            if self.timesync:
                                self.timesync.quit()
                            print(" ({}.{})".format(x, y))
                        else:
                            print(">({}.{})".format(x, y))
                            if pos[0] >= x or pos[1] >= y:
                                pos = None
                                print("invalid region selected")
                                continue
                            pos2 =(pos[0], pos[1], x - pos[0], y - pos[1])
                            pos = None
                            if self.timesync:
                                self.timesync.quit()
                            self.timesync = TimeSync(*pos2)
                            self.timesync.start()
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            user32.UnregisterHotKey(None, 1)


class TimeSync(threading.Thread):

    def __init__(self, left, top, width, height):
        super(TimeSync, self).__init__()
        w, h = pyautogui.size()
        assert left < w and (left + width) < w
        assert height < h and (height + height) < h
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.should_quit = False

    def run(self):
        last_second = None
        while not self.should_quit:
            ts1 = datetime.now()
            img = pyautogui.screenshot(region=(self.left, self.top, self.width, self.height))
            s = pytesseract.image_to_string(img)
            mobj = re.search(r"(\d+)\D+(\d{2})\D+(\d{2})", s)
            if mobj:
                #overhead = int((datetime.now() - ts1).microseconds/1000)
                hour = int(mobj.group(1))
                minute = int(mobj.group(2))
                second = int(mobj.group(3))

                if last_second is None:
                    last_second = second
                if last_second != second:
                    last_second = second

                    print("Sys %s UI %02d:%02d Diff %d.%d"%(ts1.strftime("%M:%S.%f"), minute, second, (ts1.second - second), (ts1.microsecond/1000)))
            else:
                last_second = None
                print("-->{}<--".format(s))

    def quit(self):
        self.should_quit = True
        self.join()

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
