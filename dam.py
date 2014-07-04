#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from PySide.QtCore import *
from PySide.QtGui import *
from PySide.QtDeclarative import QDeclarativeView

from pprint import pprint

import win32com
import win32com.client

import pymongo
from pymongo import MongoClient
from bson.objectid import ObjectId

client = MongoClient('localhost', 27017)

db = client.test_database
posts = db.posts
posts.drop()

import datetime
post = {}



class Main(QWidget):
    def __init__(self):
        self.borderless = True
        super(Main, self).__init__()
        self.initUI()

    def initUI(self):
        self.list = QListWidget(self)
        uiWidth = 800
        padding = 50
        self.setGeometry(0, 0, uiWidth, 260)
        # Create Qt application and the QDeclarative view
        if self.borderless:
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

        # ######## User Input *QLineEdit*
        fileSearch = QLineEdit(self)
        fileSearch.setFont(QFont('SansSerif', 20))
        fileSearch.setGeometry(0, 0, uiWidth-padding, 25)
        fsWidth = (uiWidth - fileSearch.width()) / 2
        fileSearch.setGeometry(fsWidth, 10, uiWidth-padding, 35)
        self.list.setGeometry(fsWidth, 50, uiWidth-padding, 200)
        fileSearch.textChanged.connect(self.onChanged)

        self.move((QApplication.desktop().availableGeometry().center().x() - self.rect().center().x()), 0)
        label = QLabel()
        label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        label.setAlignment(Qt.AlignBottom | Qt.AlignLeft)

        sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
        ns = sh.NameSpace(r'C:\Dropbox\External\Video Tutorials')
        column = 0
        columns = []
        while True:
            colname = ns.GetDetailsOf(None, column)
            if not colname:
                break
            columns.append(colname)
            column += 1

        for item in ns.Items():
            label.setText(label.text() + "\n")
            o = ObjectId()
            for column in range(len(columns)):
                colval = ns.GetDetailsOf(item, column)
                if colval:
                    post[columns[column]] = colval
                    post['_id'] = o
                    label.setText(label.text() + "\n" + columns[column] + " | " + colval)
            pprint(post)
            post_id = posts.insert(post)

        # MainFrameLayout.addWidget(label)

        self.show()
        print(post_id)
        # Enter Qt main loop

    def onChanged(self, text):
        self.list.clear()
        items = list(posts.find({"Name":{"$regex": text}}))
        # pprint(items)
        i = 0
        for item in items:

            test = "test"
            listItem = QListWidgetItem()
            listItem.setText(item['Name'])
            self.list.addItem(listItem)
            print(item['Name'])
            i+=1


def main():
    app = QApplication(sys.argv)
    main = Main()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()