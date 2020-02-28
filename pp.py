import ctypes
import os
import re
import threading
import time
from configparser import ConfigParser
from ctypes import wintypes
from datetime import datetime, timedelta

import pyautogui
import pytesseract
from win32 import win32gui
from win32.lib import win32con

IMG_RECG_CONFIDENCE = 0.8

class HotKey(threading.Thread):

    def __init__(self):
        super(HotKey, self).__init__()
        self.index = 99
        self.keys = {}

    def run(self):
        user32 = ctypes.windll.user32
        for k in self.keys:
            info = self.keys[k]
            if not user32.RegisterHotKey(None, k, info[0], info[1]):
                raise RuntimeError
        try:
            msg = wintypes.MSG()
            while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    if msg.wParam in self.keys:
                        info = self.keys[msg.wParam]
                        info[2](info[3])
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            for _id in self.keys:
                user32.UnregisterHotKey(None, _id)

    def register(self, mod, key, handler, arg):
        self.keys[self.index] = (mod, key, handler, arg)
        self.index += 1


class MyConfig(ConfigParser):

    SECTION_GENERAL = "general"
    SECTION_POSITION = "position"
    SECTION_TIME = "time"
    FORMAT_TIME = "%M:%S.%f"
    FORMAT_TIME_DELTA = "%S.%f"
    
    def __init__(self, filename):
        self.cfg = ConfigParser()
        if os.path.exists(filename):
            self.cfg.read(filename)
        self.lock = threading.Lock()
        self.filename = filename

    def __update(self):
        with open(self.filename, "w") as f:
            self.cfg.write(f)

    def dump(self):
        with self.lock:
            for s in self.cfg:
                section = self.cfg[s]
                for k in section:
                    print(s, k, section[k])

    def get_position(self, key):
        with self.lock:
            vals = self.cfg[MyConfig.SECTION_POSITION][key].split(",")
            return (int(vals[0]), int(vals[1]))

    def set_position(self, key, x, y):
        with self.lock:
            self.cfg[MyConfig.SECTION_POSITION][key] = "%d,%d"%(x, y)
            self.__update()

    def get_time(self, key):
        with self.lock:
            t = datetime.strptime(self.cfg[MyConfig.SECTION_TIME][key], MyConfig.FORMAT_TIME)
            return datetime.now().replace(minute=t.minute, second=t.second, microsecond=t.microsecond)

    def set_time(self, key, time):
        with self.lock:
            self.cfg[MyConfig.SECTION_TIME][key] = time.strftime(MyConfig.FORMAT_TIME)
            self.__update()

    def get_time_adjust(self):
        with self.lock:
            t_adj = self.cfg.getfloat(MyConfig.SECTION_TIME, "time_adjust")
            print("time_adjust {}".format(t_adj))
            return timedelta(seconds=int(t_adj), microseconds=int((t_adj -int(t_adj)) * 1e6))

    def get_top_most(self):
        with self.lock:
            return self.cfg.getboolean(MyConfig.SECTION_GENERAL, "top_most")

    def get_working_area(self):
        with self.lock:
            (x, y, w, h) = self.cfg.get(MyConfig.SECTION_GENERAL, "working_area").split(",")
            return (int(x), int(y), int(w), int(h))

    def get_click_input(self):
        with self.lock:
            return self.cfg.getboolean(MyConfig.SECTION_GENERAL, "click_input")


class Clicker(threading.Thread):

    def __init__(self, config, immediate):
        super(Clicker, self).__init__()
        self.cfg = config
        self.immediate = immediate

    def run(self):
        if self.immediate:
            t1 = datetime.now()
            t2 = t1 + timedelta(seconds=7)
        else:
            t_adj = self.cfg.get_time_adjust()
            t1 = self.cfg.get_time("time_submit") + t_adj
            t2 = self.cfg.get_time("time_upload") + t_adj
        
        print("ts1: {} ts2: {}".format(t1.strftime(MyConfig.FORMAT_TIME), t2.strftime(MyConfig.FORMAT_TIME)))

        while datetime.now() < t1:
            pass

        print("add price")
        if self.immediate:
            pyautogui.click(*self.cfg.get_position("add_300"))
        else:
            pyautogui.click(*self.cfg.get_position("add_cust"))

        print("bid")
        pyautogui.click(*self.cfg.get_position("bid"))

        for i in range(3):
            pos = self.locate_img_pos("nin_de.png")
            if pos:
                print("title found ({})".format(i))
                pos = self.locate_img_pos("refresh.png")
                if pos:
                    print("refresh")
                    pyautogui.click(*pos)
                break
        else:
            print("no title found")
        
        if self.cfg.get_click_input():
            pyautogui.click(*self.cfg.get_position("input_window"))

        submit_calibrate = self.locate_img_pos("submit.png")
        if submit_calibrate:
            print("submit pos re-calibrate {},{}".format(*submit_calibrate))

        while datetime.now() < t2:
            pass

        print("submit")
        if submit_calibrate:
            pyautogui.click(*submit_calibrate)
        else:
            p = self.cfg.get_position("submit")
            pyautogui.click(*p)
            pyautogui.click(p[0] - 90, p[1])
            pyautogui.click(p[0] + 90, p[1])

        print("done!")

    def locate_img_pos(self, path_to_img):
        pos = pyautogui.locateOnScreen(path_to_img, confidence=IMG_RECG_CONFIDENCE, region=self.cfg.get_working_area())
        if pos:
            center = pyautogui.center(pos)
            return (center.x, center.y)
        else:
            return None


class Calibration(threading.Thread):

    def __init__(self, config):
        super(Calibration, self).__init__()
        self.cfg = config

    def locate_img_pos(self, path_to_img):
        pos = pyautogui.locateOnScreen(path_to_img, confidence=IMG_RECG_CONFIDENCE, region=self.cfg.get_working_area())
        if pos:
            center = pyautogui.center(pos)
            return (center.x, center.y)
        else:
            return None

    def check(self, path):
        name = os.path.splitext(os.path.basename(path))[0]
        pos = self.locate_img_pos(path)
        if pos:
            print("{} {},{}".format(name, *pos))
            self.cfg.set_position(name, *pos)
        else:
            print("NO {}".format(name))

    def run(self):
        print("=" * 60)
        self.check("add_cust.png")
        self.check("add_300.png")
        self.check("bid.png")
        self.check("submit.png")
        self.check("refresh.png")
        print("calibration done!")
    

def TopMostMe(x, y, width, height):
    hwnd = win32gui.GetForegroundWindow()
    #(left, top, right, bottom) = win32gui.GetWindowRect(hwnd)
    #win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, left, top, right-left, bottom-top, 0)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, x, y, width, height, 0)

def hk_pos(cfg):
    x, y = pyautogui.position()
    print("position {},{}".format(x, y))

def hk_bid_immediate(cfg):
    Clicker(cfg, True).start()

def hk_bid_final(cfg):
    if False:
        now = datetime.now()
        cfg.set_time("time_submit", now)
        cfg.set_time("time_upload", now + timedelta(seconds=7))
    Clicker(cfg, False).start()

def hk_calibrate(cfg):
    Calibration(cfg).start()

def hk_debug(cfg):
    print("debug")
    box = pyautogui.locateOnScreen("submit.png", confidence=IMG_RECG_CONFIDENCE)
    print(box)

if __name__=='__main__':
    cfg = MyConfig("settings.ini")
    
    if cfg.get_top_most():
        TopMostMe(0, 0, 500, 200)

    cfg.dump()

    hotkey = HotKey()
    hotkey.register(win32con.MOD_CONTROL, ord("0"), hk_bid_immediate, cfg)
    hotkey.register(win32con.MOD_CONTROL, ord("6"), hk_bid_final, cfg)
    hotkey.register(win32con.MOD_CONTROL, ord("1"), hk_pos, None)
    hotkey.register(win32con.MOD_CONTROL, ord("2"), hk_calibrate, cfg)
    hotkey.register(win32con.MOD_CONTROL, ord("4"), hk_debug, cfg)
    hotkey.start()

    hotkey.join()
