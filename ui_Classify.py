# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'D:\Icarus\learning\VS-codes\Jupyter\明日方舟抽卡分析\断点更新\寻访分类.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(960, 637)
        self.gridLayout_2 = QtWidgets.QGridLayout(Form)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.NoteBTN = QtWidgets.QPushButton(Form)
        self.NoteBTN.setObjectName("NoteBTN")
        self.gridLayout.addWidget(self.NoteBTN, 0, 0, 1, 1)
        self.RemoveBTN = QtWidgets.QPushButton(Form)
        self.RemoveBTN.setObjectName("RemoveBTN")
        self.gridLayout.addWidget(self.RemoveBTN, 0, 1, 1, 1)
        self.WorkLayout = QtWidgets.QHBoxLayout()
        self.WorkLayout.setObjectName("WorkLayout")
        self.gridLayout.addLayout(self.WorkLayout, 1, 0, 1, 2)
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.gridLayout.addWidget(self.tableWidget, 2, 0, 1, 2)
        self.gridLayout_2.addLayout(self.gridLayout, 0, 0, 1, 1)

        self.retranslateUi(Form)
        self.NoteBTN.clicked.connect(Form.NoteSearch)
        self.RemoveBTN.clicked.connect(Form.RemoveNote)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.NoteBTN.setText(_translate("Form", "寻访标注"))
        self.RemoveBTN.setText(_translate("Form", "标注移除"))