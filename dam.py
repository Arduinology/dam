#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from PySide.QtCore import *
from PySide.QtGui import *
from PySide.QtDeclarative import QDeclarativeView

import win32com
import win32com.client

# Create Qt application and the QDeclarative view
app = QApplication(sys.argv)
MainWindow = QMainWindow(parent=None)
MainFrame = QFrame(MainWindow)
MainFrame.setFrameStyle(QFrame.NoFrame)
MainWindow.setCentralWidget(MainFrame)
MainFrameLayout = QVBoxLayout(MainFrame)

view = QDeclarativeView()
# Create an URL to the QML file
url = QUrl('views/view.qml')
# Set the QML file and show
view.setSource(url)
label = QLabel()
label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
label.setAlignment(Qt.AlignBottom | Qt.AlignLeft)

sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
ns = sh.NameSpace(r'C:\Dropbox\External\Video Tutorials')
column = 0
columns = []
while True:
    colname=ns.GetDetailsOf(None, column)
    if not colname:
        break
    columns.append(colname)
    column += 1

for item in ns.Items():
    label.setText(label.text()+"\n")
    for column in range(len(columns)):
        colval=ns.GetDetailsOf(item, column)
        if colval:
            label.setText(label.text()+"\n"+columns[column]+" | "+colval)


MainFrameLayout.addWidget(label)
MainWindow.show()
# Enter Qt main loop
sys.exit(app.exec_())