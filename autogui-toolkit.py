#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: eswizardry
# @Date:   2016-02-08 18:13:11
# @Last Modified by:   eswizardry
# @Last Modified time: 2016-02-14 11:51:50
"""

All coordinates assume a screen resolution of 1366x768, and Chrome
maximized with the Bookmarks Toolbar enabled.
Down key has been hit 3 times to center play area in browser.
x_pad = 219
y_pad = 238
Play area =  x_pad+1, y_pad+1, 888, 718
"""
import sys
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

from PIL import ImageGrab
from PIL import ImageOps
from numpy import *
import os
import time
import zlib
import win32api
import win32con
import win32ui
import win32gui
import pyautogui
# Globals
# ------------------
window_name = 'Calculator'
VK_CODE = {'0': 0x30,
           '1': 0x31,
           '2': 0x32,
           '3': 0x33,
           '4': 0x34,
           '5': 0x35,
           '6': 0x36,
           '7': 0x37,
           '8': 0x38,
           '9': 0x39,
           'a': 0x41,
           'b': 0x42,
           'c': 0x43,
           'd': 0x44,
           'e': 0x45,
           'f': 0x46,
           'g': 0x47,
           'h': 0x48,
           'i': 0x49,
           'j': 0x4A,
           'k': 0x4B,
           'l': 0x4C,
           'm': 0x4D,
           'n': 0x4E,
           'o': 0x4F,
           'p': 0x50,
           'q': 0x51,
           'r': 0x52,
           's': 0x53,
           't': 0x54,
           'u': 0x55,
           'v': 0x56,
           'w': 0x57,
           'x': 0x58,
           'y': 0x59,
           'z': 0x5A,
           'F1': 0x70,
           'F2': 0x71,
           'F3': 0x72,
           'F4': 0x73,
           'F5': 0x74,
           'F6': 0x75,
           'F7': 0x76,
           'F8': 0x77,
           'F9': 0x78,
           'F10': 0x79,
           'F11': 0x7A,
           'F12': 0x7B,
           'F13': 0x7C,
           'F14': 0x7D,
           'F15': 0x7E,
           'F16': 0x7F,
           'F17': 0x80,
           'F18': 0x81,
           'F19': 0x82,
           'F20': 0x83,
           'F21': 0x84,
           'F22': 0x85,
           'F23': 0x86,
           'F24': 0x87}


def get_target_window(resize=False):
    hwnd = win32gui.FindWindow(None, window_name)
    if hwnd:
        if resize:
            win32gui.MoveWindow(hwnd, screen_start_x, screen_start_y, x_size, y_size, True)
        else:
            xleft, ytop, xright, ybottom = win32gui.GetWindowRect(hwnd)
            width = xright - xleft
            height = ybottom - ytop
            return xleft, ytop, width, height
    else:
        w = QWidget()
        QMessageBox.critical(w, "Error", window_name+" instance is not exist.")
        exit()


def get_cords():
    xleft, ytop, width, height = get_target_window()
    x, y = win32api.GetCursorPos()
    x = x - xleft
    y = y - ytop
    pos = (x, y)
    return pos


def get_pos_pixel(cord):
    im = grab_screen()

    xleft, ytop, width, height = get_target_window()
    if cord < (width, height):
        pos = (cord[0], cord[1])
        pixel = im.getpixel(pos)
        print("Pixel color(R,G,B):  ", pixel)
    else:
        print("Mouse position is out of target window")


def grab_screen():
    xleft, ytop, width, height = get_target_window()
    box = (xleft, ytop, xleft + width, ytop + height)
    image = ImageGrab.grab(box)
    return image


def get_mouse_info():
    keyPress = 0
    pos = 0, 0
    oldPos = 0, 0
    while keyPress == 0:
        keyPress = win32api.GetAsyncKeyState(VK_CODE['q'])
        try:
            pos = get_cords()
            xleft, ytop, width, height = get_target_window()
        except Exception:
            pass
        else:
            pass
        if oldPos != pos:
            oldPos = pos
            print('======================================')
            print('Window origin: ' + str((xleft, ytop)))
            print('Window size: ' + str((width, height)))
            print("Mouse Position(x,y): ", pos)
            try:
                get_pos_pixel(pos)
            except Exception:
                pass
            else:
                pass


def capture_target_window():
    xleft, ytop, width, height = get_target_window()
    box = (xleft, ytop, xleft + width, ytop + height)
    im = ImageGrab.grab(box)
    im.save(os.getcwd() + "\\snap\\" + '\\' + window_name + '__' + str(int(time.time())) + '.png', 'PNG')
    print('Target window capture is done.')


def capture_all_screen():
    box = (0, 0, 1366, 768)
    im = ImageGrab.grab(box)
    im.save(os.getcwd() + "\\snap\\" + '\\' + window_name + '__' + str(int(time.time())) + '.png', 'PNG')
    print('All screen capture is done.')


class Autogui_tk_ui(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        QToolTip.setFont(QFont('SansSerif', 10))

        self.setToolTip('This is a <b>QWidget</b> widget')

        self.btn = QPushButton('Mouse', self)
        self.btn.clicked.connect(get_mouse_info)
        self.btn.setToolTip('This is a <b>QPushButton</b> widget')
        self.btn.resize(self.btn.sizeHint())
        self.btn.move(10, 10)
        self.btn2 = QPushButton('Capture', self)
        self.btn2.clicked.connect(capture_target_window)
        self.btn2.setToolTip('This is a <b>QPushButton</b> widget')
        self.btn2.resize(self.btn.sizeHint())
        self.btn2.move(100, 10)
        self.lbl1 = QLabel(self)
        self.lbl1.setText('Target window name: ')
        self.lbl1.move(200, 10)
        self.ledit1 = QLineEdit(self)
        self.ledit1.setText(window_name)
        self.ledit1.textChanged[str].connect(self.onChanged)
        self.ledit1.move(200, 25)

        self.setGeometry(900, 100, 400, 50)
        self.setWindowTitle('autoGui Tool kit')
        self.show()

    def onChanged(self, text):
        try:
            global window_name
            window_name = self.ledit1.text()
        except Exception:
            QMessageBox.about(self, 'Error', 'No config file')
            pass
        else:
            pass
        finally:
            pass


if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Autogui_tk_ui()
    sys.exit(app.exec_())
