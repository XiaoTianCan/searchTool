# -*- coding:utf-8 -*-
import sys
import os
import myApp
from PyQt5.QtWidgets import QApplication

import os
from  docx import Document
from  docx.shared import  Pt
from  docx.oxml.ns import  qn
import PyQt5.sip
#from PyQt5 import QtGui, QtCore, uic
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import datetime

if __name__ == '__main__':
    if not os.path.exists("database"):
        os.makedirs("database")
    app = QApplication(sys.argv)
    mainWindow = myApp.MyApp()
    addWindow = myApp.MyAdd()
    editWindow = myApp.MyEdit()
    setWindow = myApp.MySet()
    typeWindow = myApp.MyType()
    mainWindow.addWin(addWindow, editWindow, setWindow)
    addWindow.addWin(mainWindow)
    editWindow.addWin(mainWindow)
    setWindow.addWin(typeWindow, mainWindow)
    typeWindow.addWin(setWindow)
    mainWindow.show()
    sys.exit(app.exec_())


#input check: space, length, blank, limitation of input number
#other bugs
#adjust layout for proper input size limit

#cd D:\searchProj2
#pyinstaller -F -c -i search.ico main.py
#pyinstaller -F -w -i search.ico main.py