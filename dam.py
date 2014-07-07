#!/usr/bin/env python
# -*- coding: utf-8 -*-

from pprint import pprint

from whoosh.index import create_in
from whoosh.fields import *
from whoosh.qparser import QueryParser
from whoosh.analysis import NgramAnalyzer
import win32com
import win32com.client
from pymongo import MongoClient
from bson.objectid import ObjectId

from PySide.QtCore import *
from PySide.QtGui import *


# Whoosh
schema = Schema(Filename=TEXT(stored=True, analyzer=NgramAnalyzer(1)),
                File_description=TEXT(stored=True, analyzer=NgramAnalyzer(1)),
                Date_created=TEXT(stored=True, analyzer=NgramAnalyzer(1)))
ix = create_in("indexdir", schema)
writer = ix.writer()


# Database
client = MongoClient('localhost', 27017)
db = client.test_database
posts = db.posts
posts.drop()

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
        fileSearch.setGeometry(0, 0, uiWidth - padding, 25)
        fsWidth = (uiWidth - fileSearch.width()) / 2
        fileSearch.setGeometry(fsWidth, 10, uiWidth - padding, 35)
        self.list.setGeometry(fsWidth, 50, uiWidth - padding, 200)
        fileSearch.textChanged.connect(self.onChanged)

        self.move((QApplication.desktop().availableGeometry().center().x() - self.rect().center().x()), 0)
        label = QLabel()
        label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        label.setAlignment(Qt.AlignBottom | Qt.AlignLeft)

        sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
        ns = sh.NameSpace(r'C:')
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
            if 'File description' not in post:
                post['File description'] = ''
            if 'Date created' not in post:
                post['Date created'] = ''
            writer.add_document(Filename=post['Filename'],
                                File_description=post['File description'],
                                Date_created=post['Date created'])
            pprint(post['Filename'])
            post_id = posts.insert(post)

        writer.commit()

        # MainFrameLayout.addWidget(label)

        self.show()
        print(post_id)
        # Enter Qt main loop

    def onChanged(self, text):
        self.list.clear()
        items = list(posts.find({"Name": {"$regex": text}}))
        with ix.searcher() as searcher:
            query = QueryParser("Filename", ix.schema).parse(text)
            results = searcher.search(query)
            for result in results:
                listItem = QListWidgetItem()
                listItem.setText(result['Filename'])
                self.list.addItem(listItem)
                self.createListItem(self)

    def createListItem(self, qwidget):
        print("a")
        delegate = QItemDelegate()


class resultDelegate(QItemDelegate):
    def __init__(self, parent = None):
        QItemDelegate.__init__(self, parent)

    def createEditor(self, parent, option, index):
        editor = QSpinBox(parent)
        editor.setMinimum(0)
        editor.setMaximum(5)
        editor.installEventFilter(self)

        return editor

    def setEditorData(self, spinBox, index):
        value = index.model().data(index, Qt.DisplayRole)
        spinBox.setValue(value)

    def setModelData(self, spinBox, model, index):
        spinBox.interpretText()
        value = spinBox.value()

        model.setData(index, value)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

def main():
    app = QApplication(sys.argv)

    model = QStandardItemModel(4, 2)
    model.setHorizontalHeaderLabels(["Icon", "File Info"])
    tableView = QTableView()
    tableView.setModel(model)

    delegate = resultDelegate()
    tableView.setItemDelegate(delegate)

    for row in range (4):
        for column in range(2):
            index = model.index(row, column, QModelIndex())
            model.setData(index, (row+1) * (column+1))

    tableView.setWindowTitle("test")
    tableView.show()
    # main = Main()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()