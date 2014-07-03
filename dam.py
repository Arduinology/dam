#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from PySide.QtCore import *
from PySide.QtGui import *
from PySide.QtDeclarative import QDeclarativeView

import win32com
import win32com.client

import pymongo
from pymongo import MongoClient
client = MongoClient('localhost', 27017)
db = client.test_database
collection = db.test_collection

class Main(QWidget):
    def __init__(self):
        self.borderless = True
        super(Main, self).__init__()
        self.initUI()

    def initUI(self):
        uiWidth = 800
        self.setGeometry(0,0,uiWidth,55)
        # Create Qt application and the QDeclarative view
        if self.borderless == True:
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

        ######### User Input *QLineEdit*
        fileSearch = QLineEdit(self)
        fileSearch.setFont(QFont('SansSerif', 20))
        fileSearch.setGeometry(0,0, 770, 25)
        fsWidth = (uiWidth - fileSearch.width())/2
        fileSearch.setGeometry(fsWidth,10, 770, 35)
        fileSearch.textChanged.connect(self.onChanged)

        self.move((QApplication.desktop().availableGeometry().center().x() - self.rect().center().x()),0)
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


        # MainFrameLayout.addWidget(label)
        print label.text()
        self.show()
        # Enter Qt main loop

    def onChanged(self, text):
        print(text)

def main():
    app = QApplication(sys.argv)
    main = Main()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()