#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys

from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QAction, QApplication, QDesktopWidget, QGridLayout, QHBoxLayout, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPushButton, QVBoxLayout, qApp)


class YJK(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 设置按键
        btn = QPushButton('生成')
        qbtn = QPushButton('Quit')
        # 布局
        titleEdit = self.Buju(btn, qbtn)
        # 提示框
        titleEdit.textChanged[str].connect(self.onChanged)
        # 退出提醒
        qbtn.clicked.connect(QCoreApplication.instance().quit)
        qbtn.resize(qbtn.sizeHint())
        qbtn.move(350, 400)
        # 菜单栏
        folderAction = self.FolderAction()
        exitAction = self.ExitAction()

        self.statusBar()

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&文件')
        fileMenu.addAction(folderAction)
        fileMenu.addAction(exitAction)

        # 总界面
        self.resize(550, 550)
        self.center()
        self.setWindowTitle('YJK处理')
        # 显示
        self.show()

    def ExitAction(self):
        exitAction = QAction(QIcon('exit.png'), '&退出', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('退出程序')
        exitAction.triggered.connect(qApp.quit)
        return exitAction

    def FolderAction(self):
        folderAction = QAction(QIcon('folder.png'), '&文件夹', self)
        folderAction.setShortcut('Ctrl+W')
        folderAction.setStatusTip('打开文件夹')
        return folderAction

    def Buju(self, btn, qbtn):
        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(btn)
        hbox.addWidget(qbtn)

        self.title = QLabel('地址')
        titleEdit = QLineEdit()
        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.title, 0, 0)
        grid.addWidget(titleEdit, 1, 0)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(grid)
        vbox.addLayout(hbox)

        self.setLayout(vbox)
        return titleEdit

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message', "Are you sure to quit?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def onChanged(self, text):

        self.title.setText(text)
        self.title.adjustSize()

    def go(self):
        self.initUI()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    yjk = YJK()
    yjk.go()
    sys.exit(app.exec_())
