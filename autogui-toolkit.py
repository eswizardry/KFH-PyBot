#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: eswizardry
# @Date:   2016-02-08 18:13:11
# @Last Modified by:   Bancha Rajainthong
# @Last Modified time: 2016-10-31 21:51:46
"""

All coordinates assume a screen resolution of 1366x768, and Chrome
maximized with the Bookmarks Toolbar enabled.
Down key has been hit 3 times to center play area in browser.
x_pad = 219
y_pad = 238
Play area =  x_pad+1, y_pad+1, 888, 718
"""
import os
import sys
import time
import zlib

import win32api
import win32gui

from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

from numpy import *
from PIL import ImageGrab
from PIL import ImageOps

# Globals
# ------------------
global_window_name = 'Droid4X 0.10.4 Beta'
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


def getTargetWindow(resize=False):
    hwnd = win32gui.FindWindow(None, global_window_name)
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
        QMessageBox.critical(w, "Error", global_window_name+" instance is not exist.")
        exit()


def getCords():
    xleft, ytop, width, height = getTargetWindow()
    x, y = win32api.GetCursorPos()
    x = x - xleft
    y = y - ytop
    pos = (x, y)
    return pos


def getPosPixel(cord):
    im = grabScreen()

    xleft, ytop, width, height = getTargetWindow()
    if cord < (width, height):
        pos = (cord[0], cord[1])
        pixel = im.getpixel(pos)
        print("Pixel color(R,G,B):  ", pixel)
    else:
        print("Mouse position is out of target window")


def grabScreen():
    xleft, ytop, width, height = getTargetWindow()
    box = (xleft, ytop, xleft + width, ytop + height)
    image = ImageGrab.grab(box)
    return image


def getMouseInfo():
    key_press = 0
    pos = 0, 0
    old_pos = 0, 0
    while key_press == 0:
        QApplication.processEvents()
        key_press = win32api.GetAsyncKeyState(VK_CODE['q'])
        try:
            pos = getCords()
            xleft, ytop, width, height = getTargetWindow()
        except Exception:
            pass

        if old_pos != pos:
            old_pos = pos
            print('======================================')
            print('Window origin: ' + str((xleft, ytop)))
            print('Window size: ' + str((width, height)))
            print("Mouse Position(x,y): ", pos)
            try:
                getPosPixel(pos)
            except Exception:
                pass


def snapTargetWindow():
    xleft, ytop, width, height = getTargetWindow()
    box = (xleft, ytop, xleft + width, ytop + height)
    im = ImageGrab.grab(box)
    im.save(os.getcwd() + "\\snap\\" + '\\' + global_window_name + '__' + str(int(time.time())) + '.png', 'PNG')
    print('Target window capture is done.')


def snapEntireWindow():
    box = (0, 0, 1366, 768)
    im = ImageGrab.grab(box)
    im.save(os.getcwd() + "\\snap\\" + '\\' + global_window_name + '__' + str(int(time.time())) + '.png', 'PNG')
    print('All screen capture is done.')


def getKFHWindow(resize=False):
    hwnd = win32gui.FindWindow(None, "Droid4X 0.10.4 Beta")
    if hwnd:
        if resize:
            win32gui.MoveWindow(hwnd, screen_start_x, screen_start_y, x_size, y_size, True)
        else:
            xleft, ytop, xright, ybottom = win32gui.GetWindowRect(hwnd)
            return xleft-1, ytop-1
    else:
        w = QWidget()
        QMessageBox.critical(w, "Error", "No Droid4X 0.9.0 Beta instance is exist.")
        exit()


def imgGrab(x_start, y_start, x_end, y_end):
    xleft, ytop = getKFHWindow()
    box = (xleft+x_start, ytop+y_start, xleft+x_end, ytop+y_end)
    im = ImageGrab.grab(box)
    return im


def getSkill():
    im = ImageOps.grayscale(imgGrab(570, 90, 624, 144))
    master = array(im.getcolors())
    master = zlib.crc32(master)
    print(master)
    return master


def getEnemy():
    # Enemy X3 stage
    # im = ImageOps.grayscale(imgGrab(167, 200, 266, 305))

    # Enemy X2 stage
    im = ImageOps.grayscale(imgGrab(301, 200, 400, 305))

    enemy = array(im.getcolors())
    enemy = zlib.crc32(enemy)
    # enemy = enemy.sum()
    print(enemy)
    return enemy


class AutoGuiTk(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        QToolTip.setFont(QFont('SansSerif', 10))

        self.mouse_track_btn = QPushButton('Mouse Track', self)
        self.mouse_track_btn.clicked.connect(getMouseInfo)
        self.mouse_track_btn.setToolTip('To get <b>Mouse</b> information')
        self.mouse_track_btn.resize(self.mouse_track_btn.sizeHint())

        self.snap_target_btn = QPushButton('Snap window', self)
        self.snap_target_btn.clicked.connect(snapTargetWindow)
        self.snap_target_btn.setToolTip('To <b>snap</b> targeting window')
        self.snap_target_btn.resize(self.snap_target_btn.sizeHint())

        self.snap_all_btn = QPushButton('Snap All', self)
        self.snap_all_btn.clicked.connect(snapEntireWindow)
        self.snap_all_btn.setToolTip('To <b>snap</b> entire window')
        self.snap_all_btn.resize(self.snap_all_btn.sizeHint())

        self.get_skill_btn = QPushButton('Get Skill', self)
        self.get_skill_btn.clicked.connect(getSkill)
        self.get_skill_btn.setToolTip('To <b>get</b> skill')
        self.get_skill_btn.resize(self.get_skill_btn.sizeHint())

        self.get_enemy_btn = QPushButton('Get Enemy', self)
        self.get_enemy_btn.clicked.connect(getEnemy)
        self.get_enemy_btn.setToolTip('To <b>get</b> enemy')
        self.get_enemy_btn.resize(self.get_enemy_btn.sizeHint())

        self.input_lbl = QLabel(self)
        self.input_lbl.setText('Enter Window name: ')
        self.input_qle = QLineEdit(self)
        self.input_qle.setText(global_window_name)
        self.input_qle.textChanged[str].connect(self.onChanged)

        #  Layout
        self.h_box1 = QHBoxLayout()
        self.h_box2 = QHBoxLayout()
        self.h_box1.addWidget(self.mouse_track_btn)
        self.h_box1.addWidget(self.snap_target_btn)
        self.h_box1.addWidget(self.snap_all_btn)
        self.h_box1.addWidget(self.get_skill_btn)
        self.h_box1.addWidget(self.get_enemy_btn)
        self.h_box1.addStretch(1)
        self.h_box2.addWidget(self.input_lbl)
        self.h_box2.addWidget(self.input_qle)
        self.h_box2.addStretch(1)
        # Main horizontal layout
        self.main_vbox = QVBoxLayout()
        self.main_vbox.addLayout(self.h_box1)
        self.main_vbox.addLayout(self.h_box2)
        # self.main_vbox.addStretch(1)

        self.setLayout(self.main_vbox)
        self.setWindowTitle('Auto Gui TK')
        self.show()

    def onChanged(self, text):
        try:
            global global_window_name
            global_window_name = self.input_qle.text()
        except Exception:
            QMessageBox.about(self, 'Error', 'Invalid input value!!!')


if __name__ == '__main__':

    app = QApplication(sys.argv)
    gui_tk = AutoGuiTk()
    sys.exit(app.exec_())
