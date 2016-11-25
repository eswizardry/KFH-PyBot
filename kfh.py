#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: eswizardry
# @Date:   2015-10-02 21:20:46
# @Last Modified by:   Bancha Rajainthong
# @Last Modified time: 2016-11-11 20:52:12
"""
KFH PyBot V 0.1
"""
import os
import sys
import time
import zlib
import win32api
import win32con
import win32gui
import datetime
import pyautogui
import pyHook
import pythoncom

from PIL import ImageGrab
from PIL import ImageOps
from numpy import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *


# Globals
# ------------------
x_size = 700
y_size = 500
mouse_click_pos = 0, 0
screen_start_x = 150
screen_start_y = 100

ENEMY_X3 = 0
ENEMY_X2 = 1

YANG_XIAO     = 0
XIE_XUN       = 1
GOLDEN_HONG   = 2
SECRET_OLDMAN = 3
LI_XUNHUAN    = 4
# DONGFANG      = 0
# XIMEN         = 0
# YAO_YUE       = 0
# YI_DENG       = 0


class Enemy:

    """docstring for ClassName"""
    x3 = [1104817058, 3439856729, 2995251515, 1023797234, 2828370573]
    x2 = [895109948, 1236643701, 1394344544, 2680447795, 1809947768]

SKILLS_DICT = {
    'green_critical_skip':   1523289932,
    'green_anticri_skip':    3783338200,
    'green_dodge_skip':      3998951483,
    'green_hit_skip':      1897901501,
    'green_meditate_skip':  1897901501,
    'green_speed_skip':      2234121853,  #
    'green_inner_skip':    990064379,  #
    'green_protect_skip':    1135142825,
    'green_ATK_skip':         3936627135,
    'green_HP_skip':         2423802541,  #

    'blue_critical_blood':  2427144802,
    'blue_anticri':    1436898714,
    'blue_meditate_blood':    2349639666,
    'blue_dodge':    3182278704,
    'blue_hit_blood':    4293102806,
    'blue_inner_skip':    2316652322,
    'blue_protect_skip':    950597022,
    'blue_speed_skip':    2737793844,
    'blue_ATK':      1005769789,
    'blue_HP':       2815223053,

    'violet_anticri_blood':  139314889,
    'violet_dodge_blood':    946822975,
    'violet_hit_blood':    1281857028,
    'violet_speed':  2204472528,
    'violet_inner_blood':  325248999,
    'violet_protect':  3031685518,
    'violet_ATK_blood':    1005769789,
    'vilolet_HP_blood':  2699097725,

    # แม่เฒ่า
    'special_leela_blood':  684369648,
    # ลี้ชิวจุ้ย
    'special_increase_damage_blood':  2559285603,
    # อู๋หย๋าจือ
    'special_crumble_rate_blood':  325248999,
    # อั้งชิกกง
    'special_crumble_level_blood':  2349639666,
    # เอี้ยก้วย
    'special_prodjood_blood':  3678131005,
    # เหล่งนึ่ง
    # 'special_reduce_damage_blood':  4293102806,

}


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


def screenGrab():
    xleft, ytop = getKFHWindow()
    box = (xleft+1, ytop+1, xleft+x_size, ytop+y_size)
    image = ImageGrab.grab(box)
    return image


def getEnemy(enemyStage):
    if enemyStage == ENEMY_X3:
        im = ImageOps.grayscale(imgGrab(167, 200, 266, 305))
    else:
        im = ImageOps.grayscale(imgGrab(301, 200, 400, 305))
    enemy = array(im.getcolors())
    enemy = zlib.crc32(enemy)
    # enemy = enemy.sum()
    return enemy


def getSkill():
    im = ImageOps.grayscale(imgGrab(570, 90, 624, 144))
    master = array(im.getcolors())
    master = zlib.crc32(master)
    return master


def invisibleClick(cord):
    xleft, ytop = getKFHWindow()
    xcur, ycur = pyautogui.position()
    pyautogui.click((xleft + cord[0], ytop + cord[1]))
    pyautogui.moveTo(xcur, ycur)


def get_cords():
    xleft, ytop = getKFHWindow()
    x, y = win32api.GetCursorPos()
    x = x - xleft
    y = y - ytop
    pos = (x, y)
    return pos


def getPosPixel(cord):
    im = screenGrab()

    if cord < (600, 400):
        pos = (cord[0], cord[1])
        pixel = im.getpixel(pos)
        print("Pixel color(R,G,B):  ", pixel)
    else:
        print("Mouse position is out of target window")


def getMouseInfo():
    keyPress = 0
    pos = 0, 0
    oldPos = 0, 0
    while keyPress == 0:
        keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)
        try:
            pos = get_cords()
        except Exception:
            pass
        else:
            pass
        if oldPos != pos:
            oldPos = pos
            print("Mouse Position(x,y): ", pos)
            # pyautogui.moveTo(pos)
            try:
                getPosPixel(pos)
            except Exception:
                pass
            else:
                pass
            # print pixel
            # invisibleClick()


def skipD4XUpdate():
    # Ignore D4XUpdate
    IgnoreD4XUpdatePos = (505, 392)
    time.sleep(.1)
    invisibleClick(IgnoreD4XUpdatePos)


def resizeD4X():
    # Move and Resize D4x window
    getKFHWindow(True)


def resizeConsoleWindow():
    cslhwnd = win32gui.FindWindow(None, "pyBot")
    if cslhwnd:
        win32gui.MoveWindow(cslhwnd, 860, 530, 500, 230, True)


class KFHPyBot(QMainWindow):
    buffLimit = [3000, 3000, 1000, 1000, 0]
    buffCurrent = [0, 0, 0, 0, 0]
    # Buf Enum
    HP  = 0
    PWR = 1
    PRT = 2
    INT = 3
    AGI = 4

    gold_limit = 5

    # Stage Enum
    StageX1 = 0
    StageX2 = 1
    StageX3 = 2

    stageCurrent = 1
    stageLimit = [240, 350, 500]
    avoidLimitx3 = [0, 0, 0, 0, 0]
    avoidLimitx2 = [0, 0, 0, 0, 0]

    # Buff Position
    getBuff_30_BTNPos = (484, 216)
    getBuff_15_BTNPos = (376, 216)
    getBuff_3_BTNPos = (260, 216)

    HP30   = (213, 78, 153)
    PWR30  = (205, 70, 34)
    PRT30  = (244, 204, 75)
    INT30  = (26, 145, 149)
    AGI30  = (21, 165, 55)

    HP15   = (208, 86, 150)
    PWR15  = (204, 80, 52)
    PRT15  = (239, 199, 81)
    INT15  = (40, 143, 146)
    AGI15  = (38, 159, 65)

    HP3    = (221, 101, 165)
    PWR3   = (213, 93, 61)
    PRT3   = (251, 211, 99)
    INT3   = (50, 160, 163)
    AGI3   = (46, 176, 75)

    def __init__(self):
        super().__init__()
        self.initUI()
        # skipD4XUpdate()
        resizeD4X()
        resizeConsoleWindow()
        # self.enteringKFH()
        self.guirestore((os.getcwd() + "\\saves\\" + "default", ".ini"))

    def resetStat(self):
        self.buffCurrent = [0, 0, 0, 0, 0]
        self.stageCurrent = 1
        self.updateBuff2GUI()

    def showSaveDialog(self):
        fname = QFileDialog.getSaveFileName(self, 'Save file', os.getcwd() + "\\saves\\", ".ini")
        self.guisave(fname)

    def showLoadDialog(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', os.getcwd() + "\\saves\\")
        self.guirestore(fname)

    def initUI(self):
        # Tool bar and status bar
        resizeAction = QAction(QIcon('rsc\\d4x-icon.png'), 'ปรับขนาด D4X', self)
        resizeAction.setShortcut('Ctrl+Z')
        resizeAction.setStatusTip('ปรับขนาดหน้าจอ Droid4X')
        resizeAction.triggered.connect(resizeD4X)

        attackEvilTowerAction = QAction(QIcon('rsc\\kfh.png'), 'บุกหอมาร', self)
        attackEvilTowerAction.setShortcut('Ctrl+A')
        attackEvilTowerAction.setStatusTip('บุกหอมาร')
        attackEvilTowerAction.triggered.connect(self.attackEvilParty)

        clearStatAction = QAction(QIcon('rsc\\reset-icon.png'), 'รีเซตค่าหอมาร', self)
        clearStatAction.setShortcut('Ctrl+R')
        clearStatAction.setStatusTip('รีเซตค่าสถานะของหอมาร')
        clearStatAction.triggered.connect(self.resetStat)

        saveAction = QAction(QIcon('rsc\\save-icon.png'), 'บันทึกค่าหอมาร', self)
        saveAction.setShortcut('Ctrl+S')
        saveAction.setStatusTip('บันทึกค่าหอมาร')
        saveAction.triggered.connect(self.showSaveDialog)

        loadAction = QAction(QIcon('rsc\\load-icon.png'), 'โหลดค่าหอมาร', self)
        loadAction.setShortcut('Ctrl+L')
        loadAction.setStatusTip('โหลดค่าหอมาร')
        loadAction.triggered.connect(self.showLoadDialog)

        trainingAction = QAction(QIcon('rsc\\training-icon.png'), 'ศิษย์ฝึก', self)
        trainingAction.setShortcut('Ctrl+A')
        trainingAction.setStatusTip('ศิษย์ฝึก')
        trainingAction.triggered.connect(self.training)

        brushAction = QAction(QIcon('rsc\\brush-icon.png'), 'ย่อยพู่กัน', self)
        brushAction.setShortcut('Ctrl+B')
        brushAction.setStatusTip('ย่อยพู่กัน')
        brushAction.triggered.connect(self.extractBrush)

        mouseInfoAction = QAction(QIcon('rsc\\mouse-icon.png'), 'Mouse Info', self)
        mouseInfoAction.setShortcut('Ctrl+M')
        mouseInfoAction.setStatusTip('Mouse Info')
        mouseInfoAction.triggered.connect(getMouseInfo)

        exitAction = QAction(QIcon('rsc\\exit-icon.png'), 'ออกจากโปรแกรม', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('ออกจากโปรแกรม')
        exitAction.triggered.connect(self.close)

        self.statusBar()

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(exitAction)

        toolbar = self.addToolBar('Exit')

        toolbar.addAction(resizeAction)
        toolbar.addAction(attackEvilTowerAction)
        toolbar.addAction(clearStatAction)
        toolbar.addAction(saveAction)
        toolbar.addAction(loadAction)
        toolbar.addAction(trainingAction)
        toolbar.addAction(brushAction)
        toolbar.addAction(mouseInfoAction)
        toolbar.addAction(exitAction)

        # ===========================================================

        # Stage current
        self.lblStage = QLabel(self)
        self.lblStage.setText('ด่านปัจจุบัน: ')
        self.qleStage = QLineEdit(self)
        self.qleStage.setText('0001')
        self.qleStage.textChanged[str].connect(self.onChanged)

        # Stage Limit
        self.lblStageLimit = QLabel(self)
        self.lblStageLimit.setText('จำกัดด่าน: ')
        # x3
        self.qleStageLimit3x = QLineEdit(self)
        self.qleStageLimit3x.setText('1000')
        self.lblStageLimit3x = QLabel(self)
        self.lblStageLimit3x.setText('ด่าน x3')
        self.qleStageLimit3x.textChanged[str].connect(self.onChanged)
        # x2
        self.qleStageLimit2x = QLineEdit(self)
        self.qleStageLimit2x.setText('5000')
        self.lblStageLimit2x = QLabel(self)
        self.lblStageLimit2x.setText('ด่าน x2')
        self.qleStageLimit2x.textChanged[str].connect(self.onChanged)
        # x1
        self.qleStageLimit1x = QLineEdit(self)
        self.qleStageLimit1x.setText('9999')
        self.lblStageLimit1x = QLabel(self)
        self.lblStageLimit1x.setText('ด่าน x1')
        self.qleStageLimit1x.textChanged[str].connect(self.onChanged)

        # Stage Layout
        stageLimit_hbox1 = QHBoxLayout()
        stageLimit_hbox1.addWidget(self.qleStageLimit3x)
        stageLimit_hbox1.addWidget(self.lblStageLimit3x)
        stageLimit_hbox1.addStretch(1)
        stageLimit_hbox2 = QHBoxLayout()
        stageLimit_hbox2.addWidget(self.qleStageLimit2x)
        stageLimit_hbox2.addWidget(self.lblStageLimit2x)
        stageLimit_hbox2.addStretch(1)
        stageLimit_hbox3 = QHBoxLayout()
        stageLimit_hbox3.addWidget(self.qleStageLimit1x)
        stageLimit_hbox3.addWidget(self.lblStageLimit1x)
        stageLimit_hbox3.addStretch(1)

        stageLimit_vbox = QVBoxLayout()
        stageLimit_vbox.addLayout(stageLimit_hbox1)
        stageLimit_vbox.addLayout(stageLimit_hbox2)
        stageLimit_vbox.addLayout(stageLimit_hbox3)
        stageLimit_vbox.addStretch(1)

        stage_vbox = QVBoxLayout()
        stage_vbox.addWidget(self.lblStage)
        stage_vbox.addWidget(self.qleStage)
        stage_vbox.addWidget(self.lblStageLimit)
        stage_vbox.addLayout(stageLimit_vbox)
        stage_vbox.addStretch(1)

        # Enemy avoiding
        enemy = ['เอี้ยเซียว', 'เจี่ยซุ่น', 'กิมฮ้ง', 'แป๊ะม่อ', 'ลี้คิมฮวง']
        length = len(enemy)
        self.enemy_cbx3 = []
        self.enemy_cbx2 = []
        self.enemy_qlex3 = []
        self.enemy_qlex2 = []
        self.inner_enemy_hbox = []

        self.avoid_lbl1 = QLabel(self)
        self.avoid_lbl1.setText('หลบด่าน x3:  ')
        self.avoid_lbl2 = QLabel(self)
        self.avoid_lbl2.setText('หลบด่าน x2:  ')
        self.avoid_hbox = QHBoxLayout()
        self.avoid_hbox.addSpacing(30)
        self.avoid_hbox.addWidget(self.avoid_lbl1)
        self.avoid_hbox.addWidget(self.avoid_lbl2)
        self.avoid_hbox.addStretch(1)

        # enemy_cbx3
        for i in enemy:
            self.enemy_qlex3.append(QLineEdit(self))
            self.enemy_cbx3.append(QCheckBox(' ', self))
            self.enemy_qlex2.append(QLineEdit(self))
            self.enemy_cbx2.append(QCheckBox(i, self))
            self.inner_enemy_hbox.append(QHBoxLayout())
        for i in range(0, length):
            self.inner_enemy_hbox[i].addSpacing(30)
            self.inner_enemy_hbox[i].addWidget(self.enemy_qlex3[i])
            self.inner_enemy_hbox[i].addWidget(self.enemy_cbx3[i])
            self.inner_enemy_hbox[i].addWidget(self.enemy_qlex2[i])
            self.inner_enemy_hbox[i].addWidget(self.enemy_cbx2[i])
            self.inner_enemy_hbox[i].addStretch()
            self.enemy_qlex3[i].setFixedSize(35, 15)
            self.enemy_qlex3[i].setText('000')
            self.enemy_qlex3[i].textChanged[str].connect(self.onChanged)
            self.enemy_qlex2[i].setFixedSize(35, 15)
            self.enemy_qlex2[i].setText('000')
            self.enemy_qlex2[i].textChanged[str].connect(self.onChanged)

        self.enemy_vbox = QVBoxLayout()
        self.enemy_vbox.addLayout(self.avoid_hbox)
        for i in range(0, length):
            self.enemy_vbox.addLayout(self.inner_enemy_hbox[i])
        self.enemy_vbox.addStretch(1)

        stage_hbox = QHBoxLayout()
        stage_hbox.addLayout(stage_vbox)
        stage_hbox.addStretch(1)
        stage_hbox.addLayout(self.enemy_vbox)

        # Buff
        self.lblBuff = QLabel(self)
        self.lblBuff.setText('บัฟหอมาร: ')
        # HP
        self.lblBuffHP = QLabel(self)
        self.lblBuffHP.setText('เลือด:')
        self.lblBuffCurrentHP = QLabel(self)
        self.lblBuffCurrentHP.setText('0000 / ')
        self.qleBuffHP = QLineEdit(self)
        self.qleBuffHP.setText(str(self.buffLimit[0]))
        self.qleBuffHP.textChanged[str].connect(self.onChanged)
        # Power
        self.lblBuffPWR = QLabel(self)
        self.lblBuffPWR.setText('กำลัง:')
        self.lblBuffCurrentPWR = QLabel(self)
        self.lblBuffCurrentPWR.setText('0000 / ')
        self.qleBuffPWR = QLineEdit(self)
        self.qleBuffPWR.setText(str(self.buffLimit[1]))
        self.qleBuffPWR.textChanged[str].connect(self.onChanged)
        # Protect
        self.lblBuffPRT = QLabel(self)
        self.lblBuffPRT.setText('ป้องกัน:')
        self.lblBuffCurrentPRT = QLabel(self)
        self.lblBuffCurrentPRT.setText('0000 / ')
        self.qleBuffPRT = QLineEdit(self)
        self.qleBuffPRT.setText(str(self.buffLimit[2]))
        self.qleBuffPRT.textChanged[str].connect(self.onChanged)
        # Inner force
        self.lblBuffINT = QLabel(self)
        self.lblBuffINT.setText('กลภน:')
        self.lblBuffCurrentINT = QLabel(self)
        self.lblBuffCurrentINT.setText('0000 / ')
        self.qleBuffINT = QLineEdit(self)
        self.qleBuffINT.setText(str(self.buffLimit[3]))
        self.qleBuffINT.textChanged[str].connect(self.onChanged)
        # AGI
        self.lblBuffAGI = QLabel(self)
        self.lblBuffAGI.setText('ท่าร่าง:')
        self.lblBuffCurrentAGI = QLabel(self)
        self.lblBuffCurrentAGI.setText('0000 / ')
        self.qleBuffAGI = QLineEdit(self)
        self.qleBuffAGI.setText(str(self.buffLimit[4]))
        self.qleBuffAGI.textChanged[str].connect(self.onChanged)
        # Star use
        self.lblUseStar = QLabel(self)
        self.lblUseStar.setText('ใช้ดาวไป:')
        self.qleUseStar = QLineEdit(self)
        self.qleUseStar.setText('00000')
        self.qleUseStar.textChanged[str].connect(self.onChanged)

        # layout inner H1
        buff_hbox1 = QHBoxLayout()
        buff_hbox1.addWidget(self.lblBuffHP)
        buff_hbox1.addWidget(self.lblBuffCurrentHP)
        buff_hbox1.addWidget(self.qleBuffHP)
        buff_hbox1.addStretch(1)
        # layout inner H2
        buff_hbox2 = QHBoxLayout()
        buff_hbox2.addWidget(self.lblBuffPWR)
        buff_hbox2.addWidget(self.lblBuffCurrentPWR)
        buff_hbox2.addWidget(self.qleBuffPWR)
        buff_hbox2.addStretch(1)
        # layout inner H3
        buff_hbox3 = QHBoxLayout()
        buff_hbox3.addWidget(self.lblBuffPRT)
        buff_hbox3.addWidget(self.lblBuffCurrentPRT)
        buff_hbox3.addWidget(self.qleBuffPRT)
        buff_hbox3.addStretch(1)
        # layout inner H4
        buff_hbox4 = QHBoxLayout()
        buff_hbox4.addWidget(self.lblBuffINT)
        buff_hbox4.addWidget(self.lblBuffCurrentINT)
        buff_hbox4.addWidget(self.qleBuffINT)
        buff_hbox4.addStretch(1)
        # layout inner H5
        buff_hbox5 = QHBoxLayout()
        buff_hbox5.addWidget(self.lblBuffAGI)
        buff_hbox5.addWidget(self.lblBuffCurrentAGI)
        buff_hbox5.addWidget(self.qleBuffAGI)
        buff_hbox5.addStretch(1)
        # layout inner H6
        buff_hbox6 = QHBoxLayout()
        buff_hbox6.addWidget(self.lblUseStar)
        buff_hbox6.addWidget(self.qleUseStar)
        buff_hbox6.addStretch(1)
        # layout mid
        buff_leftvbox = QVBoxLayout()
        buff_leftvbox.addLayout(buff_hbox1)
        buff_leftvbox.addLayout(buff_hbox2)
        buff_leftvbox.addLayout(buff_hbox3)
        buff_leftvbox.addStretch(1)
        buff_rightvbox = QVBoxLayout()
        buff_rightvbox.addLayout(buff_hbox4)
        buff_rightvbox.addLayout(buff_hbox5)
        buff_rightvbox.addLayout(buff_hbox6)
        buff_rightvbox.addStretch(1)
        # layout outerH2
        buff_hboxo2 = QHBoxLayout()
        buff_hboxo2.addLayout(buff_leftvbox)
        buff_hboxo2.addLayout(buff_rightvbox)
        buff_hboxo2.addStretch(1)

        # layout inner H1
        buff_hboxo1 = QHBoxLayout()
        buff_hboxo1.addWidget(self.lblBuff)
        buff_hboxo1.addStretch(1)
        # layout outerV
        buff_vbox = QVBoxLayout()
        buff_vbox.addLayout(stage_hbox)
        buff_vbox.addLayout(buff_hboxo1)
        buff_vbox.addLayout(buff_hboxo2)
        buff_vbox.addStretch(1)
        # self.setLayout(buff_vbox)

        window = QWidget()
        # Frame for layout and widget
        hbox = QHBoxLayout(window)

        self.battleOfHeroGuiAtBottomLeft()
        self.legendaryWarriorGuiAtBottomMid()
        self.climbingMtGuiAtBottomRight()

        self.midframe = QFrame(self)
        self.midframe.setFrameShape(QFrame.StyledPanel)
        self.midframe.setLayout(buff_vbox)

        # topframe = QFrame(self)
        # topframe.setFrameShape(QFrame.StyledPanel)
        # topframe.setLayout(qbtn_vbox)

        splitter1 = QSplitter(Qt.Horizontal)
        splitter1.addWidget(self.bottomleft)
        splitter1.addWidget(self.bottommid)
        splitter1.addWidget(self.bottomright)

        splitter2 = QSplitter(Qt.Vertical)
        splitter2.addWidget(self.midframe)
        splitter2.addWidget(splitter1)

        hbox.addWidget(splitter2)
        window.setLayout(hbox)

        # Set QWidget as the central layout of the main window
        self.setCentralWidget(window)

        # Main window
        self.setGeometry(1000, 30, 250, 250)
        # self.guirestore()
        self.setWindowTitle('KFH PyBot V 0.1')
        self.setWindowIcon(QIcon('rsc\\kfh.png'))

        self.show()

    def extractBrush(self):
        keyPress = 0

        while keyPress == 0:
            QApplication.processEvents()
            # Exit program when Key press = 'SPACE Bar'
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)

            # Take screen shot
            im = screenGrab()

            gotoBigBrother = (60, 112)
            pixel = im.getpixel(gotoBigBrother)
            if pixel == (12, 35, 51):
                invisibleClick(gotoBigBrother)

            gotoSecreteRoom = (380, 65)
            pixel = im.getpixel(gotoSecreteRoom)
            if pixel == (102, 66, 17):
                invisibleClick(gotoSecreteRoom)

            gotoConsiderRoom = (246, 354)
            pixel = im.getpixel(gotoConsiderRoom)
            if pixel == (25, 85, 76):
                invisibleClick(gotoConsiderRoom)

            # Conduct brush BUG
            dragAtMid = (550, 243)
            pixel = im.getpixel(dragAtMid)

            xleft, ytop = getKFHWindow()
            dragAtMid = (xleft+550, ytop+243)
            if pixel == (17, 57, 68):
                # Up4
                for i in range(0, 4):
                    pyautogui.moveTo((dragAtMid))
                    pyautogui.dragRel(0, -50, duration=0.1)

                # Down 2, Up 2 for 3 rounds
                for i in range(0, 3):
                    # Down 2
                    pyautogui.moveTo((dragAtMid))
                    pyautogui.dragRel(0, 50, duration=0.1)
                    pyautogui.moveTo((dragAtMid))
                    pyautogui.dragRel(0, 50, duration=0.1)
                    # Up 2
                    pyautogui.moveTo((dragAtMid))
                    pyautogui.dragRel(0, -50, duration=0.1)
                    pyautogui.moveTo((dragAtMid))
                    pyautogui.dragRel(0, -50, duration=0.1)
                # Up 1
                pyautogui.moveTo((dragAtMid))
                pyautogui.dragRel(0, -50, duration=0.1)

                # Stop when found blue item
                # time.sleep(0.5)
                # checkBlueItemPos =(415, 211)
                # pixel = im.getpixel(checkBlueItemPos)
                # if pixel == (49, 111, 211):
                pos = pyautogui.locateOnScreen('rsc\\blueitem-inspect.png')
                if pos:
                    print('End : Brush Extraction...')
                    break
                else:
                    print(pos)
                    # Check First
                    checkFirst = (590, 155)
                    invisibleClick(checkFirst)
                    # extract
                    extractItem = (265, 395)
                    invisibleClick(extractItem)
                    time.sleep(1.5)

    def legendaryWarriorGuiAtBottomMid(self):
        self.bottommid = QFrame(self)
        self.bottommid.setFrameShape(QFrame.StyledPanel)
        self.qbtn_legendary = QPushButton('ตำนานไร้พ่าย', self)
        self.qbtn_legendary.clicked.connect(self.legendaryWarrior)
        self.qbtn_legendary.resize(self.qbtn_legendary.sizeHint())
        self.legendary_hbox1 = QHBoxLayout()
        self.legendary_hbox1.addWidget(self.qbtn_legendary)
        self.legendary_hbox1.addStretch(1)

        self.lbl_battleLegendary11 = QLabel(self)
        self.lbl_battleLegendary11.setText('รอบต่อสู้: ')
        self.lbl_battleLegendary12 = QLabel(self)
        self.lbl_battleLegendary12.setText('00')
        self.legendary_hbox2 = QHBoxLayout()
        self.legendary_hbox2.addWidget(self.lbl_battleLegendary11)
        self.legendary_hbox2.addWidget(self.lbl_battleLegendary12)
        self.legendary_hbox2.addStretch(1)

        self.lbl_battleLegendary21 = QLabel(self)
        self.lbl_battleLegendary21.setText('ใช้ทองไป: ')
        self.lbl_battleLegendary22 = QLabel(self)
        self.lbl_battleLegendary22.setText('00 / ')
        self.qle_battleLegendary2 = QLineEdit(self)
        self.qle_battleLegendary2.setText('5')
        self.legendary_hbox3 = QHBoxLayout()
        self.legendary_hbox3.addWidget(self.lbl_battleLegendary21)
        self.legendary_hbox3.addWidget(self.lbl_battleLegendary22)
        self.legendary_hbox3.addWidget(self.qle_battleLegendary2)
        self.legendary_hbox3.addStretch(1)

        self.legendary_vbox = QVBoxLayout()
        self.legendary_vbox.addLayout(self.legendary_hbox1)
        self.legendary_vbox.addLayout(self.legendary_hbox2)
        self.legendary_vbox.addLayout(self.legendary_hbox3)
        self.legendary_vbox.addStretch(1)

        self.bottommid.setLayout(self.legendary_vbox)

    def climbingMtGuiAtBottomRight(self):
        self.bottomright = QFrame(self)
        self.bottomright.setFrameShape(QFrame.StyledPanel)
        self.qbtn_mtClimbing = QPushButton('ต่อสู้', self)
        self.qbtn_mtClimbing.clicked.connect(self.general_battle)
        self.qbtn_mtClimbing.resize(self.qbtn_mtClimbing.sizeHint())
        self.mtClimbing_hbox1 = QHBoxLayout()
        self.mtClimbing_hbox1.addWidget(self.qbtn_mtClimbing)
        self.mtClimbing_hbox1.addStretch(1)

        self.lbl_mtClimbing11 = QLabel(self)
        self.lbl_mtClimbing11.setText('รอบต่อสู้: ')
        self.lbl_mtClimbing12 = QLabel(self)
        self.lbl_mtClimbing12.setText('00')
        self.mtClimbing_hbox2 = QHBoxLayout()
        self.mtClimbing_hbox2.addWidget(self.lbl_mtClimbing11)
        self.mtClimbing_hbox2.addWidget(self.lbl_mtClimbing12)
        self.mtClimbing_hbox2.addStretch(1)

        self.lbl_mtClimbing21 = QLabel(self)
        self.lbl_mtClimbing21.setText('อันดับประลอง: ')
        self.qle_mtClimbing2 = QLineEdit(self)
        self.qle_mtClimbing2.setText('2')
        self.mtClimbing_hbox3 = QHBoxLayout()
        self.mtClimbing_hbox3.addWidget(self.lbl_mtClimbing21)
        self.mtClimbing_hbox3.addWidget(self.qle_mtClimbing2)
        self.mtClimbing_hbox3.addStretch(1)

        self.mtClimbing_vbox = QVBoxLayout()
        self.mtClimbing_vbox.addLayout(self.mtClimbing_hbox1)
        self.mtClimbing_vbox.addLayout(self.mtClimbing_hbox2)
        self.mtClimbing_vbox.addLayout(self.mtClimbing_hbox3)
        self.mtClimbing_vbox.addStretch(1)

        self.bottomright.setLayout(self.mtClimbing_vbox)

    def enteringKFH(self):
        pos = None
        while pos is None:
            pos = pyautogui.locateOnScreen('rsc\\kfh-icon.png')
        centerPos = pyautogui.center(pos)
        pyautogui.click(centerPos)

        pos = None
        while pos is None:
            pos = pyautogui.locateOnScreen('rsc\\enter-game.png')
        centerPos = pyautogui.center(pos)
        pyautogui.click(centerPos)

        pos = None
        while pos is None:
            pos = pyautogui.locateOnScreen('rsc\\enter-game2.png')
        centerPos = pyautogui.center(pos)
        pyautogui.click(centerPos)

        pos = None
        while pos is None:
            pos = pyautogui.locateOnScreen('rsc\\activity-page.png')
        centerPos = pyautogui.center(pos)
        pyautogui.click(centerPos)

    def legendaryWarrior(self):
        keyPress = 0
        battle_count = 0
        gold_usedTable = [0, 2, 5, 9, 14, 20, 27, 35, 44, 54, 65, 77, 90]
        gold_refreshCount = 0

        while keyPress == 0:
            QApplication.processEvents()
            # Exit program when Key press = 'SPACE Bar'
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)

            tm = datetime.datetime.now().time()
            print('Waiting Legendary Battle... : '+str(tm.hour)+':'+str(tm.minute)+':'+str(tm.second), end="\r")
            time.sleep(1)

            if (tm.hour == 13 or tm.hour == 21 or tm.hour == 22) and tm.minute == 0 and tm.second >= 3:
                # Eliminate Teamviewer popup
                for i in range(0, 3):  # Do 3 times for ensure
                    pos = pyautogui.locateOnScreen('rsc\\teamviewer-ok.png')
                    if pos:  # if exist click "OK"
                        centerPos = pyautogui.center(pos)
                        pyautogui.click(centerPos)

                # pos = None
                # while pos == None:
                #     pos = pyautogui.locateOnScreen('rsc\\d4x-backbutton.png')
                # centerPos = pyautogui.center(pos)
                # pyautogui.click(centerPos)

                # pos = None
                # while pos == None:
                #     pos = pyautogui.locateOnScreen('rsc\\kfh-confirmexit.png')
                # centerPos = pyautogui.center(pos)
                # pyautogui.click(centerPos)

                # self.enteringKFH()
                # time.sleep(1)
                break

        # To prevent KFH BUG show abmormal screen
        # Enter battle
        for x in range(1, 3):
            enterBattlePos = (30, 240)
            invisibleClick(enterBattlePos)

        keyPress = 0
        while keyPress == 0:
            QApplication.processEvents()
            # Exit program when Key press = 'F1'
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)

            # Take screen shot
            im = screenGrab()

            # Continue battle screen
            contBattlePos = (554, 235)
            contBattlePixel = im.getpixel(contBattlePos)
            if contBattlePixel == (213, 77, 38):
                invisibleClick(contBattlePos)

            # Enter battle
            enterBattlePos = (30, 240)
            enterBattlePixel = im.getpixel(enterBattlePos)
            if enterBattlePixel == (42, 143, 135):
                invisibleClick(enterBattlePos)

            # Join battle
            joinBattlePos = (550, 241)
            joinBattlePixel = im.getpixel(joinBattlePos)
            if joinBattlePixel == (255, 211, 109):
                invisibleClick(joinBattlePos)

            # Start battle
            startBattlePos = (590, 392)
            normalBattlePixel = (255, 220, 113)
            refreshGoldPixel = (52, 221, 238)
            startBattlePixel = im.getpixel(startBattlePos)
            if startBattlePixel == normalBattlePixel:
                battle_count += 1
                # Fire the clicks to enter battle
                for x in range(1, 10):
                    invisibleClick(startBattlePos)
                time.sleep(.1)

            # Refresh gold?
            elif startBattlePixel == refreshGoldPixel:
                # Last shot
                checkLastshotPos = (182, 116)
                lastshotPixel = im.getpixel(checkLastshotPos)
                if lastshotPixel != (192, 75, 67):
                    battle_count += 1
                    gold_refreshCount += 1
                    # Fire the clicks to enter battle
                    for x in range(1, 10):
                        invisibleClick(startBattlePos)
                    time.sleep(.1)

                else:
                    checkLessThan40PercentPos = (317, 125)
                    lessThan40PercentPixel = im.getpixel(checkLessThan40PercentPos)
                    if gold_usedTable[gold_refreshCount+1] <= self.gold_limit:
                        if lessThan40PercentPixel == (168, 168, 168):
                            battle_count += 1
                            gold_refreshCount += 1
                            # Fire the clicks to enter battle
                            for x in range(1, 10):
                                invisibleClick(startBattlePos)
                            time.sleep(.1)

            # Next
            nextBTNPos = (460, 412)
            pixel = im.getpixel(nextBTNPos)
            if pixel == (25, 75, 70):
                invisibleClick(nextBTNPos)

            # Update stats
            self.lbl_battleLegendary12.setText(str(battle_count))
            self.lbl_battleLegendary22.setText(str(gold_usedTable[gold_refreshCount]) + ' / ')

            # End
            endCheckPos = (111, 230)
            pixel = im.getpixel(endCheckPos)
            if pixel == (68, 0, 119):
                tm = datetime.datetime.now().time()
                print('End : Legendary Battle... @ '+str(tm.hour)+':'+str(tm.minute)+':'+str(tm.second), end="\r")
                break

    def matched_n_clicked(self, pos, pixel_color, img, click=True):
            get_pixel_color = img.getpixel(pos)
            if get_pixel_color == pixel_color:
                if click:
                    invisibleClick(pos)
                    time.sleep(.1)
                return True
            else:
                return False

    def get_mouse_clicked(self, img):
        global mouse_click_pos

        # Waiting for mouse click then take mouse click POS as input
        hm = pyHook.HookManager()
        hm.SubscribeMouseLeftDown(onClick)
        hm.HookMouse()

        xpos = 0
        while xpos == 0:
            QApplication.processEvents()
            pythoncom.PumpWaitingMessages()
            xpos, ypos = mouse_click_pos

        enterBattlePixel = mouse_click_pos
        # need to clear mouse_click_pos before enter next loop
        mouse_click_pos = (0, 0)
        print(enterBattlePixel)
        hm.UnhookMouse()
        get_pixel_color = img.getpixel(enterBattlePixel)
        return enterBattlePixel, get_pixel_color

    def general_battle(self):
        keyPress = 0
        battle_count = 0

        onHoldWhileChromeAccessing()

        im = screenGrab()
        enterBattlePixel, get_pixel_color = self.get_mouse_clicked(im)

        keyPress = 0
        while keyPress == 0:
            QApplication.processEvents()
            # Exit program when Key press = 'F1'
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)
            # Take screen shot
            im = screenGrab()

            # Ending dungeon battle if need gold refresh
            if self.matched_n_clicked((390, 230), (255, 203, 8), im):
                break
            elif self.matched_n_clicked((452, 238), (132, 78, 49), im):
                break

            # Start battle if color is matched
            if self.matched_n_clicked(enterBattlePixel, get_pixel_color, im):
                battle_count += 1

            # Start battle
            self.matched_n_clicked((470, 256), (246, 186, 112), im)
            # Fighting Screen
            skipFightBTNPos = (310, 400)
            if self.matched_n_clicked((365, 55), (244, 147, 34), im):
                invisibleClick(skipFightBTNPos)

            # Confirm to continue battle refresh Battle passport
            self.matched_n_clicked((375, 308), (34, 122, 105), im)

            # Next
            self.matched_n_clicked((460, 412), (27, 68, 68), im)

            # Confirm for dungeon
            self.matched_n_clicked((384, 351), (25, 88, 84), im)

            # Update stats
            self.lbl_mtClimbing12.setText(str(battle_count))

    def battleOfHeroGuiAtBottomLeft(self):
        self.bottomleft = QFrame(self)
        self.bottomleft.setFrameShape(QFrame.StyledPanel)
        self.qbtn_battle = QPushButton('ชิงเจ้ายอดยุทธ์', self)
        self.qbtn_battle.clicked.connect(self.battleOfHero)
        self.qbtn_battle.resize(self.qbtn_battle.sizeHint())
        self.battle_hbox1 = QHBoxLayout()
        self.battle_hbox1.addWidget(self.qbtn_battle)
        self.battle_hbox1.addStretch(1)

        self.lbl_battle11 = QLabel(self)
        self.lbl_battle11.setText('รอบต่อสู้: ')
        self.lbl_battle12 = QLabel(self)
        self.lbl_battle12.setText('00')
        self.battle_hbox2 = QHBoxLayout()
        self.battle_hbox2.addWidget(self.lbl_battle11)
        self.battle_hbox2.addWidget(self.lbl_battle12)
        self.battle_hbox2.addStretch(1)

        self.lbl_battle21 = QLabel(self)
        self.lbl_battle21.setText('ชนะ / แพ้: ')
        self.lbl_battle22 = QLabel(self)
        self.lbl_battle22.setText('00 / 00')
        self.battle_hbox3 = QHBoxLayout()
        self.battle_hbox3.addWidget(self.lbl_battle21)
        self.battle_hbox3.addWidget(self.lbl_battle22)
        self.battle_hbox3.addStretch(1)

        self.battle_vbox = QVBoxLayout()
        self.battle_vbox.addLayout(self.battle_hbox1)
        self.battle_vbox.addLayout(self.battle_hbox2)
        self.battle_vbox.addLayout(self.battle_hbox3)
        self.battle_vbox.addStretch(1)

        self.bottomleft.setLayout(self.battle_vbox)

    def battleOfHero(self):
        keyPress = 0
        # battle
        battle_round = 0
        battle_win = 0
        battle_lose = 0

        onHoldWhileChromeAccessing()

        while keyPress == 0:
            QApplication.processEvents()
            # Exit program when Key press = 'SPACE Bar'
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)

            # Take screen shot
            im = screenGrab()

            # Runout of ticket, Stop battle
            gotoMarketBTNPos = (460, 310)
            pixel = im.getpixel(gotoMarketBTNPos)
            if pixel == (24, 85, 78):
                break

            # Enter battle
            enterBattlePos = (47, 156)
            enterBattlePixel = im.getpixel(enterBattlePos)
            if enterBattlePixel == (250, 175, 60):
                invisibleClick(enterBattlePos)

            # Join battle
            joinBattlePos = (570, 352)
            joinBattlePixel = im.getpixel(joinBattlePos)
            if joinBattlePixel == (213, 105, 74):
                invisibleClick(joinBattlePos)

            # Start battle
            startBattlePos = (300, 386)
            startBattlePixel = im.getpixel(startBattlePos)
            if startBattlePixel == (37, 156, 161):
                # Over limit > 15 battles/day
                checkOverLimitPos = (270, 233)
                pixel = im.getpixel(checkOverLimitPos)
                if pixel == (133, 80, 51):
                    break
                else:
                    invisibleClick(startBattlePos)

            # Fighting Screen
            fightScreenPos = (365, 55)
            skipFightBTNPos = (310, 400)
            pixel = im.getpixel(fightScreenPos)
            if pixel == (244, 147, 34):
                invisibleClick(skipFightBTNPos)

            # Next
            nextBTNPos = (448, 370)
            pixel = im.getpixel(nextBTNPos)
            if pixel == (25, 102, 88):
                winCheckPos = (296, 83)
                pixel = im.getpixel(winCheckPos)
                if pixel == (74, 38, 31):
                    battle_win += 1
                    battle_round += 1
                    invisibleClick(nextBTNPos)
                else:
                    battle_lose += 1
                    battle_round += 1
                    invisibleClick(nextBTNPos)

            # Update stats
            self.lbl_battle12.setText(str(battle_round))
            self.lbl_battle22.setText(str(battle_win) + ' / ' + str(battle_lose))

            # Confirm to continue battle > 10 times
            contBTNPos = (448, 305)
            pixel = im.getpixel(contBTNPos)
            if pixel == (24, 104, 97):
                invisibleClick(contBTNPos)

            time.sleep(.1)

    def onChanged(self, text):
        try:
            self.stageCurrent = int(self.qleStage.text())
            self.stageLimit[0] = int(self.qleStageLimit3x.text())
            self.stageLimit[1] = int(self.qleStageLimit2x.text())
            self.stageLimit[2] = int(self.qleStageLimit1x.text())

            self.buffLimit[0] = int(self.qleBuffHP.text())
            self.buffLimit[1] = int(self.qleBuffPWR.text())
            self.buffLimit[2] = int(self.qleBuffPRT.text())
            self.buffLimit[3] = int(self.qleBuffINT.text())
            self.buffLimit[4] = int(self.qleBuffAGI.text())

            self.gold_limit = int(self.qle_battleLegendary2.text())

            length = len(Enemy.x3)
            for i in range(0, length):
                self.avoidLimitx3[i] = int(self.enemy_qlex3[i].text())
                self.avoidLimitx2[i] = int(self.enemy_qlex2[i].text())
        except Exception:
            QMessageBox.about(self, 'Error', 'No config file')
            pass
        else:
            pass
        finally:
            pass

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message', "Are you sure to quit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        # self.guisave()

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def getBuff(self, im):
            pixelBuff_30 = im.getpixel(self.getBuff_30_BTNPos)
            pixelBuff_15 = im.getpixel(self.getBuff_15_BTNPos)
            pixelBuff_3 = im.getpixel(self.getBuff_3_BTNPos)

            # Choose HP/PWR Buff 30%
            if self.chooseHPPWR30Buff(pixelBuff_30, pixelBuff_15):
                invisibleClick(self.getBuff_30_BTNPos)
                self.updateBuff2GUI()
            # Choose HP/PWR Buff 15%
            elif self.chooseHPPWR15Buff(pixelBuff_15):
                invisibleClick(self.getBuff_15_BTNPos)
                self.updateBuff2GUI()
            # Choose PRT/INT/AGI Buff 30%
            elif self.chooseOther30Buff(pixelBuff_30):
                invisibleClick(self.getBuff_30_BTNPos)
                self.updateBuff2GUI()
            # Choose PRT/INT/AGI Buff 15%
            elif self.chooseOther15Buff(pixelBuff_15):
                invisibleClick(self.getBuff_15_BTNPos)
                self.updateBuff2GUI()
            # Choose Buff 3%
            elif self.choose3PercentBuff(pixelBuff_3):
                    invisibleClick(self.getBuff_3_BTNPos)
                    self.updateBuff2GUI()

    def chooseHPPWR30Buff(self, pixelBuff_30, pixelBuff_15):
        # Buff nomalization - when HP/PWR greather than each other > 300%
        if (pixelBuff_30 == self.HP30) and (self.buffCurrent[self.HP] < self.buffLimit[self.HP]):
            if pixelBuff_15 == self.PWR15:
                if self.buffCurrent[self.HP] < (self.buffCurrent[self.PWR] + 300):
                    self.buffCurrent[self.HP] += 30
                    return True
            else:
                self.buffCurrent[self.HP] += 30
                return True

        if pixelBuff_30 == self.PWR30 and (self.buffCurrent[self.PWR] < self.buffLimit[self.PWR]):
            if pixelBuff_15 == self.HP15:
                if self.buffCurrent[self.PWR] < (self.buffCurrent[self.HP] + 300):
                    self.buffCurrent[self.PWR] += 30
                    return True
            else:
                self.buffCurrent[self.PWR] += 30
                return True

        return False

    def chooseHPPWR15Buff(self, pixelBuff):
        isBuffChoose = False
        if (pixelBuff == self.HP15) and (self.buffCurrent[self.HP] < self.buffLimit[self.HP]):
            self.buffCurrent[self.HP] += 15
            isBuffChoose = True
        elif pixelBuff == self.PWR15 and (self.buffCurrent[self.PWR] < self.buffLimit[self.PWR]):
            self.buffCurrent[self.PWR] += 15
            isBuffChoose = True
        else:
            pass
        return isBuffChoose

    def chooseOther30Buff(self, pixelBuff):
        isBuffChoose = False
        if pixelBuff == self.PRT30 and (self.buffCurrent[self.PRT] < self.buffLimit[self.PRT]):
            self.buffCurrent[self.PRT] += 30
            isBuffChoose = True
        elif pixelBuff == self.INT30 and (self.buffCurrent[self.INT] < self.buffLimit[self.INT]):
            self.buffCurrent[self.INT] += 30
            isBuffChoose = True
        elif pixelBuff == self.AGI30 and (self.buffCurrent[self.AGI] < self.buffLimit[self.AGI]):
            self.buffCurrent[self.AGI] += 30
            isBuffChoose = True
        else:
            pass
        return isBuffChoose

    def chooseOther15Buff(self, pixelBuff):
        isBuffChoose = False
        if pixelBuff == self.PRT15 and (self.buffCurrent[self.PRT] < self.buffLimit[self.PRT]):
            self.buffCurrent[self.PRT] += 15
            isBuffChoose = True
        elif pixelBuff == self.INT15 and (self.buffCurrent[self.INT] < self.buffLimit[self.INT]):
            self.buffCurrent[self.INT] += 15
            isBuffChoose = True
        elif pixelBuff == self.AGI15 and (self.buffCurrent[self.AGI] < self.buffLimit[self.AGI]):
            self.buffCurrent[self.AGI] += 15
            isBuffChoose = True
        else:
            pass
        return isBuffChoose

    def choose3PercentBuff(self, pixelBuff):
        isBuffChoose = False

        if pixelBuff == self.HP3:
            self.buffCurrent[self.HP] += 3
            isBuffChoose = True
        elif pixelBuff == self.PWR3:
            self.buffCurrent[self.PWR] += 3
            isBuffChoose = True
        elif pixelBuff == self.PRT3:
            self.buffCurrent[self.PRT] += 3
            isBuffChoose = True
        elif pixelBuff == self.INT3:
            self.buffCurrent[self.INT] += 3
            isBuffChoose = True
        elif pixelBuff == self.AGI3:
            self.buffCurrent[self.AGI] += 3
            isBuffChoose = True
        else:
            pass
        return isBuffChoose

    def updateBuff2GUI(self):
        # Update Buff to GUI <QTextLabel>
        self.lblBuffCurrentHP.setText(str(self.buffCurrent[self.HP]) + ' / ')
        self.lblBuffCurrentPWR.setText(str(self.buffCurrent[self.PWR]) + ' / ')
        self.lblBuffCurrentPRT.setText(str(self.buffCurrent[self.PRT]) + ' / ')
        self.lblBuffCurrentINT.setText(str(self.buffCurrent[self.INT]) + ' / ')
        self.lblBuffCurrentAGI.setText(str(self.buffCurrent[self.AGI]) + ' / ')
        # Update Stage to GUI <QTextEdit>
        self.qleStage.setText(str(self.stageCurrent))

    def isFightStage(self, enemyStage):
        fight = True
        length = len(Enemy.x3)

        if enemyStage == ENEMY_X3:
            for i in range(0, length):
                if self.enemy_cbx3[i].isChecked():
                    if self.stageCurrent > self.avoidLimitx3[i]:
                        enemy = getEnemy(enemyStage)
                        if enemy == Enemy.x3[i]:
                            fight = False
                            break
        else:
            for i in range(0, length):
                if self.enemy_cbx2[i].isChecked():
                    if self.stageCurrent > self.avoidLimitx2[i]:
                        enemy = getEnemy(ENEMY_X2)
                        if enemy == Enemy.x2[i]:
                            fight = False
                            break
        return fight

    def guisave(self, fname):
        fileName, ext = fname
        if fileName != '':
            # Strip '.ini' extension if exist.
            if '.ini' in fileName:
                fileName = fileName[:-4]

            settings = QSettings(fileName+ext, QSettings.IniFormat)
            settings.setValue("geometry", self.saveGeometry())

            settings.setValue('qleStageLimit3x', self.qleStageLimit3x.text())
            settings.setValue('qleStageLimit2x', self.qleStageLimit2x.text())
            settings.setValue('qleStageLimit1x', self.qleStageLimit1x.text())

            settings.setValue('qleBuffHP', self.qleBuffHP.text())
            settings.setValue('qleBuffPWR', self.qleBuffPWR.text())
            settings.setValue('qleBuffPRT', self.qleBuffPRT.text())
            settings.setValue('qleBuffINT', self.qleBuffINT.text())
            settings.setValue('qleBuffAGI', self.qleBuffAGI.text())

            length = len(Enemy.x3)
            for i in range(0, length):
                settings.setValue('enemy_qlex3'+str(i), self.enemy_qlex3[i].text())
                settings.setValue('enemy_qlex2'+str(i), self.enemy_qlex2[i].text())
                settings.setValue('enemy_cbx3'+str(i), self.enemy_cbx3[i].checkState())
                settings.setValue('enemy_cbx2'+str(i), self.enemy_cbx2[i].checkState())

    def guirestore(self, fname):
        fileName, ext = fname
        if fileName != '':
            settings = QSettings(fileName+ext, QSettings.IniFormat)
            self.restoreGeometry(settings.value("geometry"))

            self.qleStageLimit3x.setText(str(settings.value('qleStageLimit3x')))
            self.qleStageLimit2x.setText(str(settings.value('qleStageLimit2x')))
            self.qleStageLimit1x.setText(str(settings.value('qleStageLimit1x')))

            self.qleBuffHP.setText(str(settings.value('qleBuffHP')))
            self.qleBuffPWR.setText(str(settings.value('qleBuffPWR')))
            self.qleBuffPRT.setText(str(settings.value('qleBuffPRT')))
            self.qleBuffINT.setText(str(settings.value('qleBuffINT')))
            self.qleBuffAGI.setText(str(settings.value('qleBuffAGI')))

            length = len(Enemy.x3)
            for i in range(0, length):
                self.enemy_qlex3[i].setText(str(settings.value('enemy_qlex3'+str(i))))
                self.enemy_qlex2[i].setText(str(settings.value('enemy_qlex2'+str(i))))
                self.enemy_cbx3[i].setCheckState(int(settings.value('enemy_cbx3'+str(i))))
                self.enemy_cbx2[i].setCheckState(int(settings.value('enemy_cbx2'+str(i))))

    def attackEvilParty(self):
        keyPress = 0

        onHoldWhileChromeAccessing()

        while keyPress == 0:
            QApplication.processEvents()
            # Exit program when Key press = 'SPACE Bar'
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)

            # Take screen shot
            im = screenGrab()

            # Got dead
            checkDeadPos = (168, 131)
            deadPixel = im.getpixel(checkDeadPos)
            if deadPixel == (92, 92, 92):
                break

            # Check passport
            checkPassportPos = (580, 390)
            pixel = im.getpixel(checkPassportPos)
            if pixel == (8, 41, 48):
                invisibleClick(checkPassportPos)

            # Next
            nextBTNPos = (442, 395)
            pixel = im.getpixel(nextBTNPos)
            if pixel == (20, 93, 89):
                invisibleClick(nextBTNPos)
                self.stageCurrent += 1
                # Update Stage to GUI <QTextEdit>
                self.qleStage.setText(str(self.stageCurrent))
                # Update use star
                useStar = 0
                for x in range(0, 5):
                    useStar += self.buffCurrent[x]
                self.qleUseStar.setText(str(useStar))

            # Get Reward
            getRewardBTNPos = (370, 360)
            pixel = im.getpixel(getRewardBTNPos)
            if pixel == (29, 114, 104):
                invisibleClick(getRewardBTNPos)

            # Fighting Screen
            fightScreenPos = (365, 55)
            skipFightBTNPos = (380, 390)
            pixel = im.getpixel(fightScreenPos)
            if pixel == (244, 147, 34):
                invisibleClick(skipFightBTNPos)

            # x3BTN    [213,330],(118, 39, 15)
            # x2BTN    [347,330],(132, 42, 20)
            # x1BTN    [478,330],(157, 51, 32)
            # x3
            if (self.stageCurrent <= self.stageLimit[0]) and self.isFightStage(ENEMY_X3):
                x3BTNPos = (213, 330)
                pixel = im.getpixel(x3BTNPos)
                if pixel == (163, 94, 78):
                    invisibleClick(x3BTNPos)
            # x2
            elif self.stageCurrent <= self.stageLimit[1] and self.isFightStage(ENEMY_X2):
                x2BTNPos = (347, 330)
                pixel = im.getpixel(x2BTNPos)
                if pixel == (167, 99, 85):
                    invisibleClick(x2BTNPos)
            # x1
            elif self.stageCurrent <= self.stageLimit[2]:
                x1BTNPos = (478, 330)
                pixel = im.getpixel(x1BTNPos)
                if pixel == (167, 97, 85):
                    invisibleClick(x1BTNPos)
            else:  # Reach limit stage
                break

            # Choose Buff
            self.getBuff(im)

            time.sleep(.1)

    def isSkillExist(self, input_dict):
        skill = getSkill()
        exist = skill in input_dict.values()
        return exist

    def applyThisSkill(self):
        # Apply chicken blood skills set
        chickenblood_skills_set = {key: value for (key, value) in SKILLS_DICT.items() if '_blood' in key}
        # Normal skills set
        normal_skills_set = {key: value for (key, value) in SKILLS_DICT.items() if 'skip' not in key}

        if self.isSkillExist(chickenblood_skills_set):
            self.learningSkill(True)
            apply = True
        elif self.isSkillExist(normal_skills_set):
            self.learningSkill()
            apply = True
        else:
            apply = False
        return apply

    def learningSkill(self, use_blood=False):
        process_end = False
        while process_end is False:
            im = screenGrab()
            # apply training button
            if self.matched_n_clicked((560, 254), (254, 138, 5), im):
                print('Apply training')

            if use_blood:
                # Enabling chicken blood if not yet enable.
                self.enableChickenBlood(im)
            else:
                # Disabling chicken blood if enabled.
                self.disableChickenBlood(im)

            if self.matched_n_clicked((362, 316), (31, 130, 114), im):
                print('Apply on confirm button')

            if self.matched_n_clicked((352, 379), (64, 137, 132), im):
                print('Skip training practice')

            if self.matched_n_clicked((337, 314), (192, 181, 133), im):
                print('Continue training')
                process_end = True
            # Skip training practice screen
            # self.skipPracticeScreen()

    def skipPracticeScreen(self):
        pos = None
        while pos is None:
            pos = pyautogui.locateOnScreen('rsc\\skip-training.png')
        centerPos = pyautogui.center(pos)
        pyautogui.click(centerPos)
        print('Skip training practice')

        pos = None
        while pos is None:
            pos = pyautogui.locateOnScreen('rsc\\continue-training.png')
        centerPos = pyautogui.center(pos)
        pyautogui.click(centerPos)
        print('Continue training')

    def enableChickenBlood(self, im):
        self.matched_n_clicked((479, 221), (217, 194, 119), im)
        # print('Enable Chicken bloood')

    def disableChickenBlood(self, im):
        self.matched_n_clicked((479, 221), (33, 153, 42), im)
        # print('Disable Chicken bloood')

    def refreshSkill(self):
        im = screenGrab()
        if self.matched_n_clicked((142, 222), (132, 127, 87), im):
            print('Refresh skill')
            time.sleep(.1)

            im = screenGrab()
            if self.matched_n_clicked((392, 305), (27, 110, 99), im):
                print('Confirm to refresh violet/special skill')
            return True
        else:
            print('Training screen not exist!')
            return False

    def getCurrentSkill(self):
        skill_crc_value = getSkill()
        skill_key = [key for key, value in SKILLS_DICT.items() if value == skill_crc_value][0]
        return(skill_key)

    def training(self):
        keyPress = 0
        skill_cords = [(250, 240), (350, 240), (450, 240)]
        onHoldWhileChromeAccessing()

        while keyPress == 0:
            QApplication.processEvents()
            keyPress = win32api.GetAsyncKeyState(win32con.VK_F1)
            im = screenGrab()

            loc = pyautogui.locateOnScreen('rsc\\successful-rate-xx.png')
            # print('Successful rate < 100%!')
            # if loc is None:
            #     loc = pyautogui.locateOnScreen('rsc\\successful-rate-95.png')
            #     print('Successful rate < 100%!')

            # if loc is None:
            #     loc = pyautogui.locateOnScreen('rsc\\successful-rate-85.png')
            #     print('Successful rate < 85%!')

            if loc:
                # Stop when trainig successful rate < 100%
                print('Successful rate < 100%, stop training !!!')
                return
            else:
                # Training screen?
                if self.matched_n_clicked((142, 222), (132, 127, 87), im, False):
                    for cord in range(0, 3):
                        invisibleClick(skill_cords[cord])
                        if self.isSkillExist(SKILLS_DICT):
                            found_skill = self.getCurrentSkill()
                            print('***********************************')
                            print('Found skill: ' + found_skill)
                            self.applyThisSkill()
                        else:
                            # Stop when encounter unknow skill
                            print('Unknow skill, stop training !!!')
                            return
                    if self.refreshSkill() is False:
                        # Stop if training screen not exist
                        break
                else:
                    print('Wait for training screen !')


# Hold on when chrome remote is under accessing
def onHoldWhileChromeAccessing():
    hold_on = True
    while hold_on:
        loc = pyautogui.locateOnScreen('rsc\\stop-sharing.png')
        if loc is None:
                hold_on = False


# Mouse event detection
def onClick(event):
    global mouse_click_pos
    mouse_click_x, mouse_click_y = event.Position
    mouse_click_x = mouse_click_x - (screen_start_x - 1)
    mouse_click_y = mouse_click_y - (screen_start_y - 1)
    mouse_click_pos = (mouse_click_x, mouse_click_y)
    return True

if __name__ == '__main__':
    app = QApplication(sys.argv)
    kfhWidget = KFHPyBot()

    sys.exit(app.exec_())
