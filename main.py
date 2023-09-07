import random
import re
import threading
import time

import pyautogui
import win32api
import win32con
import win32ui
from PIL import ImageGrab
import keyboard

import win32gui
import winsound

is_crafting_enabled = True
active = False

regex_border_position = (500, 700, 501, 701)
alteration_position = win32api.MAKELONG(147, 362)

def execute():
    toplist, winlist = [], []

    def enum_cb(hwnd, results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))

    win32gui.EnumWindows(enum_cb, toplist)
    path_of_exile_application = [(hwnd, title) for hwnd, title in winlist if 'path of exile' in title.lower()]
    # just grab the hwnd for first window matching path_of_exile_application
    path_of_exile_application = path_of_exile_application[0]
    hwnd = path_of_exile_application[0]
    window = win32ui.CreateWindowFromHandle(hwnd)


    #window.SendMessage(win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, click_position)
    #time.sleep(0.1)
    #window.SendMessage(win32con.WM_LBUTTONUP, None, click_position)

    #window.SendMessage(win32con.WM_KEYDOWN, 0xA0, 0)
    while active:
        #if win32gui.GetForegroundWindow() == hwnd:
        if win32gui.GetForegroundWindow() == hwnd:
            bbox = win32gui.GetWindowRect(hwnd)
            img = ImageGrab.grab(bbox)

            if is_crafting_enabled:
                im1 = img.crop(regex_border_position)

                if len(im1.getcolors()) == 1:
                    value, rgb = im1.getcolors()[0]
                    print("R:"+str(rgb[0]) + " B:"+str(rgb[1])+ " G:"+str(rgb[2]))
                    if 235 > rgb[0] > 227 and 185 > rgb[1] > 175 and 124 > rgb[2] > 114:
                        frequency = 2500  # Set Frequency To 2500 Hertz
                        duration = 1000  # Set Duration To 1000 ms == 1 second
                        winsound.Beep(frequency, duration)
                        time.sleep(5)
                    else:
                        x_offset = random.randrange(0, 25, 1)
                        y_offset = random.randrange(0, 25, 1)
                        delay = random.randrange(100, 150, 1)
                        pyautogui.moveTo(125 + x_offset, 336 + y_offset, duration=delay/1000)
                        pyautogui.rightClick(125 + x_offset, 336 + y_offset)
                        pyautogui.moveTo(388 + x_offset * 2, 500 + y_offset, duration=delay/1000)
                        pyautogui.click(388 + x_offset * 2, 500 + y_offset)

                        pyautogui.moveTo(286 + x_offset, 415 + y_offset, duration=delay/1000)
                        pyautogui.rightClick(286 + x_offset, 415 + y_offset)
                        pyautogui.moveTo(388 + x_offset * 2, 500 + y_offset, duration=delay/1000)
                        pyautogui.click(388 + x_offset * 2, 500 + y_offset)

                        #window.SendMessage(win32con.WM_KEYUP, 0xA0, 0)


        time.sleep(0.1)



if __name__ == "__main__":
    active = not active

    f = threading.Thread(name='foreground', target=execute)

    f.start()
