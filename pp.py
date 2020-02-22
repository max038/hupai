import os
import re
import time
import threading
import pyautogui
import pytesseract
import configparser
import ctypes, win32con, ctypes.wintypes, win32gui
from datetime import datetime, timedelta

CONSOLE_LEFT = 0
CONSOLE_TOP = 0
CONSOLE_WIDTH = 500
CONSOLE_HEIGHT = 200

class HotKey(threading.Thread):

    def __init__(self, config):
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
            msg = ctypes.wintypes.MSG()
            while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    if msg.wParam in self.keys:
                        info = self.keys[msg.wParam]
                        info[2](info[3])
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            user32.UnregisterHotKey(None, 1)

    def register_hotkey(self, mod, key, handler, arg):
        self.keys[self.index] = (mod, key, handler, arg)
        self.index += 1


class MyConfig(object):

    CONFIG_FILE = "config.ini"
    SECTION_POSITION = "position"
    SECTION_TIME = "time"
    TS_FORMAT = "%M:%S.%f"
    DELTA_FORMAT = "%S.%f"

    def __init__(self):
        self.cfg = configparser.ConfigParser()
        if os.path.exists(MyConfig.CONFIG_FILE):
            self.cfg.read(MyConfig.CONFIG_FILE)
        self.lock = threading.RLock()

    def __cfg_update(self):
        with open(MyConfig.CONFIG_FILE, "w") as cfgfile:
            self.cfg.write(cfgfile)

    def set_pos(self, key, x, y):
        with self.lock:
            if MyConfig.SECTION_POSITION not in self.cfg:
                self.cfg[MyConfig.SECTION_POSITION] = {}
            section = self.cfg[MyConfig.SECTION_POSITION]
            section[key] = "%d,%d"%(x, y)
            self.__cfg_update()

    def get_pos(self, key):
        with self.lock:
            if MyConfig.SECTION_POSITION not in self.cfg:
                raise Exception("position section not exist")
            section = self.cfg[MyConfig.SECTION_POSITION]
            return [int(i) for i in section[key].split(",")]

    def dump(self):
        with self.lock:
            for s in self.cfg:
                section = self.cfg[s]
                for k in section:
                    print(s, k, section[k])

    def get_ts(self, key):
        with self.lock:
            if MyConfig.SECTION_TIME not in self.cfg:
                raise Exception("time section not exist")
            section = self.cfg[MyConfig.SECTION_TIME]
            ts = datetime.strptime(section[key], MyConfig.TS_FORMAT)
            return datetime.now().replace(minute=ts.minute, second=ts.second, microsecond=ts.microsecond)

    def set_ts(self, key, ts):
        with self.lock:
            if MyConfig.SECTION_TIME not in self.cfg:
                raise Exception("time section not exist")
            section = self.cfg[MyConfig.SECTION_TIME]
            section[key] = ts.strftime(MyConfig.TS_FORMAT)
            self.__cfg_update()

    def get_adjust(self):
        with self.lock:
            return self.cfg.getfloat("time", "adjust")


def TopMostMe():
    hwnd = win32gui.GetForegroundWindow()
    #(left, top, right, bottom) = win32gui.GetWindowRect(hwnd)
    #win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, left, top, right-left, bottom-top, 0)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, CONSOLE_LEFT, CONSOLE_TOP, CONSOLE_WIDTH, CONSOLE_HEIGHT, 0)

def hk_add(cfg):
    x, y = pyautogui.position()
    print("add {}.{}".format(x, y))
    cfg.set_pos("add", x, y)

def hk_submit(cfg):
    x, y = pyautogui.position()
    print("submit {}.{}".format(x, y))
    cfg.set_pos("submit", x, y)

def hk_ok(cfg):
    x, y = pyautogui.position()
    print("ok {}.{}".format(x, y))
    cfg.set_pos("ok", x, y)

def hk_input(cfg):
    x, y = pyautogui.position()
    print("input {}.{}".format(x, y))
    cfg.set_pos("input", x, y)

def hk_test(cfg):
    print("test")
    ts = datetime.now()

    ts1 = ts + timedelta(seconds=0)
    cfg.set_ts("test_time_submit", ts1)
    
    ts2 = ts1 + timedelta(seconds=7)
    cfg.set_ts("test_time_upload", ts2)

    cli = Clicker(cfg, True)
    cli.start()

def hk_real(cfg):
    print("real")
    cli = Clicker(cfg, False)
    cli.start()

def hk_debug(cfg):
    print("debug")
    box = pyautogui.locateOnScreen("que_ding.png", confidence=0.7)
    print(box)


class Clicker(threading.Thread):

    def __init__(self, config, isTest):
        super(Clicker, self).__init__()
        self.cfg = config
        self.isTest = isTest

    def run(self):
        t_adj = self.cfg.get_adjust()
        t_adj_delta = timedelta(seconds=int(t_adj), microseconds=((t_adj -int(t_adj)) * 1e6))

        if self.isTest:
            t1 = self.cfg.get_ts("test_time_submit")
            t2 = self.cfg.get_ts("test_time_upload")
        else:
            t1 = self.cfg.get_ts("time_submit")
            t2 = self.cfg.get_ts("time_upload")
        
        t1_adj = t1 + t_adj_delta
        t2_adj = t2 + t_adj_delta

        print("ts1: {} ts2: {} adjust: {}".format(t1.strftime(MyConfig.TS_FORMAT), t2.strftime(MyConfig.TS_FORMAT), t_adj))
        print("aj1: {} aj2: {}".format(t1_adj.strftime(MyConfig.TS_FORMAT), t2_adj.strftime(MyConfig.TS_FORMAT)))

        while datetime.now() < t1_adj:
            pass

        print("add")
        pyautogui.click(*self.cfg.get_pos("add"))
        print("submit")
        pyautogui.click(*self.cfg.get_pos("submit"))

        time.sleep(0.1)

        print("input")
        for i in range(3):
            region = (935, 526, 624, 420)
            pos = pyautogui.locateOnScreen("nin_de.png", confidence=0.8, region=region)
            if pos:
                print("title found")
                pos = pyautogui.locateOnScreen("shua_xin.png", confidence=0.8, region=region)
                if pos:
                    print("refresh")
                    center = pyautogui.center(pos)
                    pyautogui.click(center.x, center.y)
                break
        #pyautogui.click(*self.cfg.get_pos("input"))

        re_center = None
        pos = pyautogui.locateOnScreen("que_ding.png", confidence=0.7, region=region)
        if pos:
            center = pyautogui.center(pos)
            re_center = (center.x, center.y)
            print("re-center {}.{}".format(*re_center))

        while datetime.now() < t2_adj:
            pass

        print("ok")
        if re_center:
            pyautogui.click(*re_center)
        p = self.cfg.get_pos("ok")
        pyautogui.click(*p)
        pyautogui.click(p[0] - 90, p[1])
        pyautogui.click(p[0] + 90, p[1])

        print("done")


if __name__=='__main__':
    TopMostMe()
    cfg = MyConfig()
    cfg.dump()

    key = HotKey(cfg)
    key.register_hotkey(win32con.MOD_CONTROL, ord("8"), hk_add, cfg)    # add custom value
    key.register_hotkey(win32con.MOD_CONTROL, ord("9"), hk_submit, cfg) # submit custom price
    key.register_hotkey(win32con.MOD_CONTROL, ord("1"), hk_input, cfg)  # input window
    key.register_hotkey(win32con.MOD_CONTROL, ord("2"), hk_ok, cfg)     # upload final price
    key.register_hotkey(win32con.MOD_CONTROL, ord("0"), hk_test, cfg)   # trigger oneshot submit
    key.register_hotkey(win32con.MOD_CONTROL, ord("6"), hk_real, cfg)   # trigger time based submit
    key.register_hotkey(win32con.MOD_CONTROL, ord("4"), hk_debug, cfg)
    key.start()

    key.join()
