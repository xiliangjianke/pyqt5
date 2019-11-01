 # -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'F:\python\pyqt5\eric_project\TA_supplier.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!
#gridLayout_2
#数据库、表格相关
import pyodbc
import xlrd, openpyxl

#经典样式表
import qdarkstyle

import operator
import os
import io 
import sys
import time 

#用于复制、打开文件
import shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from PyQt5 import QtCore, QtGui, QtWidgets
#用于显示查询
from PyQt5.QtSql import QSqlDatabase  , QSqlQueryModel , QSqlQuery

#pyqt5连接数据库所用得到的库
from PyQt5.QtSql import QSqlTableModel, QSqlDatabase
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
#from PyQt5.QtWidgets import QFileDialog

#下拉式复选框
class CheckableComboBox(QtWidgets.QComboBox):
    def __init__(self, parent=None):
        super(CheckableComboBox, self).__init__(parent)
        self.view().pressed.connect(self.handleItemPressed)
        #设置字大小
        font = QtGui.QFont()
        font.setPointSize(12)
        self.setFont(font)
       
    #设置复选框事件
    def handleItemPressed(self, index):
        item = self.model().itemFromIndex(index)
        if item.checkState() == QtCore.Qt.Checked:
            item.setCheckState(QtCore.Qt.Unchecked)
        else:
            item.setCheckState(QtCore.Qt.Checked)

    #获取被选中选项的文本放入checkedItems
    def getCheckItem(self):
        checkedItems = []
        for index in range(self.count()):
            item = self.model().item(index)
            if item.checkState() == QtCore.Qt.Checked:
                checkedItems.append(item.text())
        return checkedItems

#将元组中的字符串转化为数字便于比较,若数据库中含有纯数字则需要此函数
def int_it(v):
    try:
        return int(v)
    except:
        return v


class Ui_Form(QMainWindow):

        #self.value = None
    def setupUi(self, Form):
        #窗口名称
        Form.setObjectName("TA信息管理系统")
        #窗口尺寸
        Form.resize(1397, 923)
        
         #字体设置
        self.font = QtGui.QFont()
        self.font.setPointSize(12)
        #首页label及layout设置
        self.gridLayout = QtWidgets.QGridLayout(Form)
        self.gridLayout.setObjectName("gridLayout")
        #创建tree
        self.treeWidget = QtWidgets.QTreeWidget(Form)
        self.treeWidget.setMinimumSize(QtCore.QSize(256, 531))
        self.treeWidget.setMaximumSize(QtCore.QSize(256, 100000))
        self.treeWidget.setObjectName("treeWidget")
        #self.treeWidget.setFont(self.font)
        #数据点监听事件，changePage函数实现页面切换
        self.treeWidget.clicked.connect(self.changePage)

        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        self.gridLayout.addWidget(self.treeWidget, 1, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setText("")
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 0, 1, 1, 1)
        self.label = QtWidgets.QLabel(Form)
        self.label.setMinimumSize(QtCore.QSize(191, 51))
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.stackedWidget = QtWidgets.QStackedWidget(Form)
        self.stackedWidget.setMinimumSize(QtCore.QSize(471, 521))
        self.stackedWidget.setObjectName("stackedWidget")
        self.page = QtWidgets.QWidget()
        self.page.setObjectName("page")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.page)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_3 = QtWidgets.QLabel(self.page)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.label_4 = QtWidgets.QLabel(self.page)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.verticalLayout.addWidget(self.label_4)
        self.label_5 = QtWidgets.QLabel(self.page)
        self.label_5.setText("")
        self.label_5.setObjectName("label_5")
        self.verticalLayout.addWidget(self.label_5)
        self.label_6 = QtWidgets.QLabel(self.page)
        self.label_6.setText("")
        self.label_6.setObjectName("label_6")
        self.verticalLayout.addWidget(self.label_6)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        self.stackedWidget.addWidget(self.page)


        

        #第二页，供应商信息录入页设置
        self.page_2 = QtWidgets.QWidget()
        self.page_2.setObjectName("page_2")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.page_2)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")

        #新增标签
        #新增零件标签
        self.label_33 = QtWidgets.QLabel(self.page_2)
        self.label_33.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_33.setObjectName("label_33")
        self.gridLayout_2.addWidget(self.label_33, 4, 3, 1, 1)
        #金属成型
        self.label_31 = QtWidgets.QLabel(self.page_2)
        self.label_31.setObjectName("label_31")
        self.gridLayout_2.addWidget(self.label_31, 4, 1, 1, 1)
        #焊接类
        self.label_32 = QtWidgets.QLabel(self.page_2)
        self.label_32.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_32.setObjectName("label_32")
        self.gridLayout_2.addWidget(self.label_32, 3, 3, 1, 1)
        #表面装饰类
        self.label_30 = QtWidgets.QLabel(self.page_2)
        self.label_30.setObjectName("label_30")
        self.gridLayout_2.addWidget(self.label_30, 3, 1, 1, 1)
        #注塑发泡类
        self.label_29 = QtWidgets.QLabel(self.page_2)
        self.label_29.setObjectName("label_29")
        self.gridLayout_2.addWidget(self.label_29, 2, 1, 1, 1)

        self.label_7 = QtWidgets.QLabel(self.page_2)
        self.label_7.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_7.setObjectName("label_7")
        self.gridLayout_2.addWidget(self.label_7, 0, 0, 1, 2)
        #供应商简称文本框，self.supAbbreviation
        self.supAbbreviation = QtWidgets.QLineEdit(self.page_2)
        self.supAbbreviation.setFont(self.font)
        self.supAbbreviation.setMaximumSize(QtCore.QSize(250, 16777215))
        self.supAbbreviation.setObjectName("supAbbreviation")
        #self.supAbbreviation.setFont(self.font)
        self.gridLayout_2.addWidget(self.supAbbreviation, 1, 2, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.page_2)
        self.label_9.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_9.setObjectName("label_9")
        self.gridLayout_2.addWidget(self.label_9, 1, 4, 1, 1)
        self.label_12 = QtWidgets.QLabel(self.page_2)
        self.label_12.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_12.setObjectName("label_12")
        self.gridLayout_2.addWidget(self.label_12, 2, 0, 1, 1)
        #工艺能力1下拉框，self.supTech1-----------------2019.7.14更新
        self.supTech1 = CheckableComboBox(self.page_2)
         #当选项被单击时获取选项文本，采用getText函数
        self.supTech1.activated.connect(self.getText)
        self.supTech1.setFont(self.font)
        self.supTech1.setMaximumSize(QtCore.QSize(200, 16777215))
        self.supTech1.setObjectName("supTech1")
        self.gridLayout_2.addWidget(self.supTech1, 2, 2, 1, 1)
        #调用函数
        self.combo_option(self.supTech1, 'injection')
        
        #工艺能力4下拉框，self.supTech4
        self.supTech4 = CheckableComboBox(self.page_2)
        self.supTech4.activated.connect(self.getText)
        self.supTech4.setFont(self.font)
        self.supTech4.setMinimumSize(QtCore.QSize(200, 0))
        self.supTech4.setMaximumSize(QtCore.QSize(111111, 16777215))
        self.supTech4.setObjectName("supTech4")
        self.gridLayout_2.addWidget(self.supTech4, 3, 4, 1, 1)
         #调用函数
        self.combo_option(self.supTech4, 'welding')

        self.label_14 = QtWidgets.QLabel(self.page_2)
        self.label_14.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_14.setObjectName("label_14")
        self.gridLayout_2.addWidget(self.label_14, 5, 0, 1, 2)
        #公司简介，self.supBrief
        self.supBrief = QtWidgets.QTextEdit(self.page_2)
        self.supBrief.setFont(self.font)
        self.supBrief.setMaximumSize(QtCore.QSize(600, 100))
        self.supBrief.setObjectName("supBrief")
        self.gridLayout_2.addWidget(self.supBrief, 5, 2, 1, 8)
        self.label_15 = QtWidgets.QLabel(self.page_2)
        self.label_15.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_15.setObjectName("label_15")
        self.gridLayout_2.addWidget(self.label_15, 6, 0, 1, 2)
        #单个插入附件，self.supOneAttach
        self.supOneAttach = QtWidgets.QPushButton(self.page_2)
        self.supOneAttach.setFont(self.font)
        self.supOneAttach.setMaximumSize(QtCore.QSize(75, 16777215))
        self.supOneAttach.setObjectName("supOneAttach")
        self.gridLayout_2.addWidget(self.supOneAttach, 6, 2, 1, 1)
        #设置函数------------------------------------------------------------
        self.supOneAttach.clicked.connect(self.get_path)
        #批量插入附件，self.supMoreAttach，暂时不需要
        #self.supMoreAttach = QtWidgets.QPushButton(self.page_2)
        #self.supMoreAttach.setObjectName("supMoreAttach")
        #self.gridLayout_2.addWidget(self.supMoreAttach, 6, 5, 1, 1)
        #模具能力标签
        self.label_23 = QtWidgets.QLabel(self.page_2)
        self.label_23.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_23.setObjectName("label_23")
        self.gridLayout_2.addWidget(self.label_23, 7, 0, 1, 2)
        #技术能力标签
        self.label_24 = QtWidgets.QLabel(self.page_2)
        self.label_24.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_24.setObjectName("label_24")
        self.gridLayout_2.addWidget(self.label_24, 7, 3, 1, 1)
        #导入数据库按钮，self.supInputDb
        self.supInputDb = QtWidgets.QPushButton(self.page_2)
        self.supInputDb.setFont(self.font)
        self.supInputDb.setMaximumSize(QtCore.QSize(150, 16777215))
        self.supInputDb.setObjectName("supInputDb")
        self.gridLayout_2.addWidget(self.supInputDb, 9, 1, 1, 2)
        #设置函数------------------------------------------------------------
        self.supInputDb.clicked.connect(self.input_into_db)
        #展示单个添加按钮，self.supOneView
        self.supOneView = QtWidgets.QPushButton(self.page_2)
        self.supOneView.setFont(self.font)
        self.supOneView.setMaximumSize(QtCore.QSize(75, 23))
        self.supOneView.setObjectName("supOneView")
        self.gridLayout_2.addWidget(self.supOneView, 9, 3, 1, 1)
        #设置函数------------------------------------------------------------
        self.supOneView.clicked.connect(self.one_input_view)

        #展示批量添加按钮， self.supMoreView
        self.supMoreView = QtWidgets.QPushButton(self.page_2)
        self.supMoreView.setFont(self.font)
        self.supMoreView.setMaximumSize(QtCore.QSize(75, 23))
        self.supMoreView.setObjectName("supMoreView")
        self.gridLayout_2.addWidget(self.supMoreView, 9, 4, 1, 2)
        #设置函数------------------------------------------------------------
        self.supMoreView.clicked.connect(self.sup_more_view)

        #清空按钮，self.supClear
        self.supClear = QtWidgets.QPushButton(self.page_2)
        self.supClear.setFont(self.font)
        self.supClear.setObjectName("supClear")
        self.gridLayout_2.addWidget(self.supClear, 9, 6, 1, 1)
        #设置函数------------------------------------------------------------
        self.supClear.clicked.connect(self.clear_all)
        #数据库备份按钮，self.supBackUp
        self.supBackUp = QtWidgets.QPushButton(self.page_2)
        self.supBackUp.setFont(self.font)
        self.supBackUp.setMaximumSize(QtCore.QSize(75, 23))
        self.supBackUp.setObjectName("supBackUp")
        self.gridLayout_2.addWidget(self.supBackUp, 9, 8, 1, 2)
        #设置函数------------------------------------------------------------
        self.supBackUp.clicked.connect(self.back_up_db)
        #供应商信息录入显示,self.subInputView
        self.subInputView = QtWidgets.QTableView(self.page_2)
        self.subInputView.setFont(self.font)
        self.subInputView.setMaximumSize(QtCore.QSize(16777215, 300))
        self.subInputView.setObjectName("subInputView")
        self.gridLayout_2.addWidget(self.subInputView, 8, 0, 1, 10)
        #工艺能力5下拉框，self.supTech5
        self.supTech5 = CheckableComboBox(self.page_2)
        self.supTech5.activated.connect(self.getText)
        self.supTech5.setFont(self.font)
        self.supTech5.setMaximumSize(QtCore.QSize(200, 16777215))
        self.supTech5.setObjectName("supTech5")
        self.gridLayout_2.addWidget(self.supTech5, 4, 4, 1, 1)
         #调用函数
        self.combo_option(self.supTech5, 'parts')

        #模具能力文本框， self.subMould
        self.supModelRearch = QtWidgets.QComboBox(self.page_2)
        self.supModelRearch.setFont(self.font)
        self.supModelRearch.setObjectName("supModelRearch")
        self.supModelRearch.addItem("")
        self.supModelRearch.setItemText(0, "")
        self.supModelRearch.addItem("")
        self.supModelRearch.addItem("")
        self.supModelRearch.addItem("")
        self.gridLayout_2.addWidget(self.supModelRearch, 7, 2, 1, 1)
        #技术能力文本框，self.subTech
        self.supTechRearch = QtWidgets.QComboBox(self.page_2)
        self.supTechRearch.setFont(self.font)
        self.supTechRearch.setObjectName("supTechRearch")
        self.supTechRearch.addItem("")
        self.supTechRearch.setItemText(0, "")
        self.supTechRearch.addItem("")
        self.supTechRearch.addItem("")
        self.supTechRearch.addItem("")
        self.gridLayout_2.addWidget(self.supTechRearch, 7, 4, 1, 3)
        #self.horizontalLayout.addLayout(self.gridLayout_2)
        self.stackedWidget.addWidget(self.page_2)
        #省文本框，self.supProvince
        self.supProvince = QtWidgets.QLineEdit(self.page_2)
        self.supProvince.setFont(self.font)
        self.supProvince.setMaximumSize(QtCore.QSize(60, 16777215))
        self.supProvince.setObjectName("supProvince")
        self.gridLayout_2.addWidget(self.supProvince, 1, 5, 1, 1)
        self.label_10 = QtWidgets.QLabel(self.page_2)
        self.label_10.setObjectName("label_10")
        self.gridLayout_2.addWidget(self.label_10, 1, 6, 1, 1)
        #市文本框，self.supCity
        self.supCity = QtWidgets.QLineEdit(self.page_2)
        self.supCity.setFont(self.font)
        self.supCity.setMaximumSize(QtCore.QSize(60, 60))
        self.supCity.setObjectName("supCity")
        self.gridLayout_2.addWidget(self.supCity, 1, 7, 1, 1)
        self.label_11 = QtWidgets.QLabel(self.page_2)
        self.label_11.setObjectName("label_11")
        self.gridLayout_2.addWidget(self.label_11, 1, 8, 1, 1)
        #供应商全称文本框，self.supName
        self.supName = QtWidgets.QLineEdit(self.page_2)
        self.supName.setFont(self.font)
        self.supName.setObjectName("supName")
        self.gridLayout_2.addWidget(self.supName, 0, 2, 1, 2)
        #工艺能力2下拉框，self.supTech2
        self.supTech2 = CheckableComboBox(self.page_2)
        self.supTech2.activated.connect(self.getText)
        self.supTech2.setFont(self.font)
        self.supTech2.setMaximumSize(QtCore.QSize(200, 16777215))
        self.supTech2.setObjectName("supTech2")
        self.gridLayout_2.addWidget(self.supTech2, 3, 2, 1, 1)
         #调用函数
        self.combo_option(self.supTech2, 'decoration')

        #工艺能力3下拉框，self.supTech2
        self.supTech3 = CheckableComboBox(self.page_2)
        self.supTech3.activated.connect(self.getText)
        self.supTech3.setFont(self.font)
        self.supTech3.setMaximumSize(QtCore.QSize(200, 16777215))
        self.supTech3.setObjectName("supTech3")
        self.gridLayout_2.addWidget(self.supTech3, 4, 2, 1, 1)
         #调用函数
        self.combo_option(self.supTech3, 'metalForming')

        self.label_13 = QtWidgets.QLabel(self.page_2)
        self.label_13.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_13.setObjectName("label_13")
        self.gridLayout_2.addWidget(self.label_13, 2, 3, 1, 1)
        #设备能力文本框，self.supEquip
        self.supEquip = QtWidgets.QLineEdit(self.page_2)
        self.supEquip.setFont(self.font)
        self.supEquip.setObjectName("supEquip")
        self.gridLayout_2.addWidget(self.supEquip, 2, 4, 1, 3)
        self.label_8 = QtWidgets.QLabel(self.page_2)
        self.label_8.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_8.setObjectName("label_8")
        self.gridLayout_2.addWidget(self.label_8, 1, 0, 1, 2)
        self.horizontalLayout.addLayout(self.gridLayout_2)
        self.stackedWidget.addWidget(self.page_2)

        #第三页，供应商信息查询页
        self.page_3 = QtWidgets.QWidget()
        self.page_3.setObjectName("page_3")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.page_3)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.gridLayout_3 = QtWidgets.QGridLayout()
        self.gridLayout_3.setHorizontalSpacing(23)
        self.gridLayout_3.setVerticalSpacing(26)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_16 = QtWidgets.QLabel(self.page_3)
        self.label_16.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_16.setObjectName("label_16")
        self.gridLayout_3.addWidget(self.label_16, 0, 0, 1, 1)
        #工艺能力查询下拉框，self.supQtech
        self.supQtech = QtWidgets.QComboBox(self.page_3)
        self.supQtech.setFont(self.font)
        self.supQtech.setMaximumSize(QtCore.QSize(200, 16777215))
        self.supQtech.setObjectName("supQtech")
        self.gridLayout_3.addWidget(self.supQtech, 0, 1, 1, 1)
        #连接查询函数
        self.supQtech.activated.connect(self.db_query)
         #调用函数
        self.combo_option(self.supQtech,'technologicalCapability')

        self.label_20 = QtWidgets.QLabel(self.page_3)
        self.label_20.setText("")
        self.label_20.setObjectName("label_20")
        self.gridLayout_3.addWidget(self.label_20, 0, 2, 1, 1)
        self.label_17 = QtWidgets.QLabel(self.page_3)
        self.label_17.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_17.setObjectName("label_17")
        self.gridLayout_3.addWidget(self.label_17, 1, 0, 1, 1)
        #供应商简称查询下拉框，self.supQbrief=================================简称下拉框
        self.supQbrief = QtWidgets.QComboBox(self.page_3)
        self.supQbrief.setFont(self.font)
        self.supQbrief.setMaximumSize(QtCore.QSize(10000000, 16777215))
        self.supQbrief.setObjectName("supQbrief")
        self.gridLayout_3.addWidget(self.supQbrief, 1, 1, 1, 1)
        #连接查询函数
        self.supQbrief.activated.connect(self.db_query)
         #调用函数
        self.combo_option(self.supQbrief, 'abbreviation')

        self.label_21 = QtWidgets.QLabel(self.page_3)
        self.label_21.setText("")
        self.label_21.setObjectName("label_21")
        self.gridLayout_3.addWidget(self.label_21, 1, 2, 1, 1)
        self.label_18 = QtWidgets.QLabel(self.page_3)
        self.label_18.setObjectName("label_18")
        self.gridLayout_3.addWidget(self.label_18, 2, 0, 1, 1)
        #供应商省查询下拉框，self.supQprovin
        self.supQprovin = QtWidgets.QComboBox(self.page_3)
        self.supQprovin.setFont(self.font)
        self.supQprovin.setMaximumSize(QtCore.QSize(1000000, 16777215))
        self.supQprovin.setObjectName("supQprovin")
        self.gridLayout_3.addWidget(self.supQprovin, 2, 1, 1, 1)
        #连接查询函数
        self.supQprovin.activated.connect(self.db_query)
         #调用函数
        self.combo_option(self.supQprovin, 'province')

        self.label_22 = QtWidgets.QLabel(self.page_3)
        self.label_22.setText("")
        self.label_22.setObjectName("label_22")
        self.gridLayout_3.addWidget(self.label_22, 2, 2, 1, 1)
        self.label_19 = QtWidgets.QLabel(self.page_3)
        self.label_19.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_19.setObjectName("label_19")
        self.gridLayout_3.addWidget(self.label_19, 3, 0, 1, 1)
        #供应商市查询下拉框，self.supQcity==============================市下拉框
        self.supQcity = QtWidgets.QComboBox(self.page_3)
        #self.supQcity.setFont(self.font)
        self.supQcity.setFont(self.font)
        self.supQcity.setMaximumSize(QtCore.QSize(10000000, 16777215))
        self.supQcity.setObjectName("supQcity")
        self.gridLayout_3.addWidget(self.supQcity, 3, 1, 1, 1)
        #连接查询函数
        self.supQcity.activated.connect(self.db_query)
         #调用函数
        self.combo_option(self.supQcity, 'city')
        #查询信息清空按钮， self.supQclear
        self.supQclear = QtWidgets.QPushButton(self.page_3)
        self.supQclear.setFont(self.font)
        self.supQclear.setMaximumSize(QtCore.QSize(75, 16777215))
        self.supQclear.setObjectName("supQclear")
        #连接函数
        self.supQclear.clicked.connect(self.clear_all)
        self.gridLayout_3.addWidget(self.supQclear, 3, 2, 1, 1)
        self.verticalLayout_3.addLayout(self.gridLayout_3)
        #采用tableview显示查询结果
        #self.supQview = QtWidgets.QTableView(self.page_3)
        #self.supQview.setMaximumSize(QtCore.QSize(16777215, 600))
        #self.supQview.setObjectName("supQview")
        #self.verticalLayout_3.addWidget(self.supQview)
        #self.stackedWidget.addWidget(self.page_3)

        #采用tablewidget显示结果，命名仍为supQview
        self.supQview = QtWidgets.QTableWidget(self.page_3)
        self.supQview.setFont(self.font)
        self.supQview.setMinimumSize(QtCore.QSize(938, 750))
        self.supQview.setObjectName("supQview")
        self.supQview.setColumnCount(0)
        self.supQview.setRowCount(0)
        self.verticalLayout_3.addWidget(self.supQview)
        self.stackedWidget.addWidget(self.page_3)

        #第四页，外协件信息录入----------------------------------------------------2019.7.20更新--------------

        

        self.page_4 = QtWidgets.QWidget()
        self.page_4.setObjectName("page_4")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.page_4)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.gridLayout_4 = QtWidgets.QGridLayout()
        self.gridLayout_4.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.gridLayout_4.setObjectName("gridLayout_4")



        #补充格式提醒label
        self.label_142 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_142.setFont(font)
        self.label_142.setObjectName("label_142")
        self.gridLayout_4.addWidget(self.label_142, 5, 3, 1, 1)
        self.label_143 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_143.setFont(font)
        self.label_143.setObjectName("label_143")
        self.gridLayout_4.addWidget(self.label_143, 7, 3, 1, 1)
        #self.horizontalLayout_2.addLayout(self.gridLayout_4)
        self.stackedWidget.addWidget(self.page_4)

        self.label_46 = QtWidgets.QLabel(self.page_4)
        self.label_46.setMinimumSize(QtCore.QSize(96, 22))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_46.setFont(font)
        self.label_46.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_46.setObjectName("label_46")
        self.gridLayout_4.addWidget(self.label_46, 7, 4, 1, 2)
        #布点供应商2
        self.exSup2 = QtWidgets.QComboBox(self.page_4)
        #更新定点供应商选项
        self.exSup2.activated.connect(lambda:self.select_sup_option(self.exSup2, 'sup'))

        #self.combo_option(self.exSup2, 'abbreviation')
        self.exSup2.setObjectName("exSup2")
        self.gridLayout_4.addWidget(self.exSup2, 10, 2, 1, 2)
        self.label_79 = QtWidgets.QLabel(self.page_4)
        self.label_79.setMaximumSize(QtCore.QSize(16777215, 50))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.label_79.setFont(font)
        self.label_79.setAlignment(QtCore.Qt.AlignCenter)
        self.label_79.setObjectName("label_79")
        self.gridLayout_4.addWidget(self.label_79, 0, 4, 1, 4)
        self.label_42 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_42.setFont(font)
        self.label_42.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_42.setObjectName("label_42")
        self.gridLayout_4.addWidget(self.label_42, 7, 8, 1, 1)
        self.exProNum = QtWidgets.QLineEdit(self.page_4)
        self.exProNum.setObjectName("exProNum")
        self.gridLayout_4.addWidget(self.exProNum, 3, 7, 1, 1)
        #布点供应商4
        self.exSup4 = QtWidgets.QComboBox(self.page_4)
        #更新定点供应商选项
        self.exSup4.activated.connect(lambda:self.select_sup_option(self.exSup4, 'sup'))

        #调用函数
        #self.combo_option(self.exSup4, 'abbreviation')

        self.exSup4.setObjectName("exSup4")
        self.gridLayout_4.addWidget(self.exSup4, 12, 2, 1, 2)
        self.label_44 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_44.setFont(font)
        self.label_44.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_44.setObjectName("label_44")
        self.gridLayout_4.addWidget(self.label_44, 6, 4, 1, 2)
        #外协件信息录入，焊接下拉框
        self.exWield = QComboBox(self.page_4)
         #当选项被单击时获取选项文本，采用getText函数
        #self.exWield.activated.connect(self.getText)
        #调用函数
        self.combo_option(self.exWield, 'welding')

        #筛选布点供应商选项
        self.exWield.activated.connect(lambda:self.exSup_option_byTech(self.exSup1))


        self.exWield.setMinimumSize(QtCore.QSize(150, 0))
        self.exWield.setMaximumSize(QtCore.QSize(176, 16777215))
        self.exWield.setObjectName("exWield")
        self.gridLayout_4.addWidget(self.exWield, 7, 9, 1, 6)
        self.label_37 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_37.setFont(font)
        self.label_37.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_37.setObjectName("label_37")
        self.label_37.setMinimumSize(QtCore.QSize(90, 0))
        self.gridLayout_4.addWidget(self.label_37, 3, 0, 1, 1)
        self.exProNumLine = QtWidgets.QLineEdit(self.page_4)

        #自动补全
        #self.auto_complete_option( 'projectNumber')

        self.completer = QCompleter(self.auto_complete_option( 'projectNumber'))
        self.completer.setFilterMode(Qt.MatchContains)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)        
        self.exProNumLine.setCompleter(self.completer)

           

        self.exProNumLine.setMinimumSize(QtCore.QSize(150, 0))
        self.exProNumLine.setMaximumSize(QtCore.QSize(176, 16777215))
        self.exProNumLine.setObjectName("exProNumLine")
        self.gridLayout_4.addWidget(self.exProNumLine, 2, 6, 1, 2)
        self.label_76 = QtWidgets.QLabel(self.page_4)
        self.label_76.setText("")
        self.label_76.setObjectName("label_76")
        self.gridLayout_4.addWidget(self.label_76, 10, 9, 1, 1)
        self.label_75 = QtWidgets.QLabel(self.page_4)
        self.label_75.setText("")
        self.label_75.setObjectName("label_75")
        self.gridLayout_4.addWidget(self.label_75, 12, 9, 1, 1)
        #布点供应商3
        self.exSup3 = QtWidgets.QComboBox(self.page_4)
        #更新定点供应商选项
        self.exSup3.activated.connect(lambda:self.select_sup_option(self.exSup3, 'sup'))
        #调用函数
        #self.combo_option(self.exSup3, 'abbreviation')

        self.exSup3.setObjectName("exSup3")
        self.gridLayout_4.addWidget(self.exSup3, 11, 2, 1, 2)
        self.exSeleSup = QtWidgets.QComboBox(self.page_4)
        #添加空选项用于初始化
        self.exSeleSup.addItem('')
        self.exSeleSup.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.exSeleSup.setObjectName("exSeleSup")
        self.gridLayout_4.addWidget(self.exSeleSup, 13, 2, 1, 2)
        self.label_86 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_86.setFont(font)
        self.label_86.setObjectName("label_86")
        self.gridLayout_4.addWidget(self.label_86, 3, 6, 1, 1)
        #布点供应商2所在地
        self.exSup2Place = QtWidgets.QComboBox(self.page_4)
        #更新定点地点选项
        self.exSup2Place.activated.connect(lambda:self.select_sup_option(self.exSup2Place, 'city'))

        #self.combo_option(self.exSup2Place, 'city')

        self.exSup2Place.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.exSup2Place.setObjectName("exSup2Place")
        self.gridLayout_4.addWidget(self.exSup2Place, 10, 5, 1, 1)
        self.label_71 = QtWidgets.QLabel(self.page_4)
        self.label_71.setText("")
        self.label_71.setObjectName("label_71")
        self.gridLayout_4.addWidget(self.label_71, 14, 0, 1, 1)
        #布点供应商1所在地
        self.exSup1Place = QtWidgets.QComboBox(self.page_4)
        #更新定点选项
        self.exSup1Place.activated.connect(lambda:self.select_sup_option(self.exSup1Place, 'city'))
        #调用函数
        self.combo_option(self.exSup1Place, 'city')
        self.exSup1Place.setMinimumSize(QtCore.QSize(80, 0))
        self.exSup1Place.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.exSup1Place.setObjectName("exSup1Place")
        self.gridLayout_4.addWidget(self.exSup1Place, 9, 5, 1, 1)
        #布点供应商3所在地
        self.exSup3Place = QtWidgets.QComboBox(self.page_4)
        #更新定点地点选项
        self.exSup3Place.activated.connect(lambda:self.select_sup_option(self.exSup3Place, 'city'))

        self.combo_option(self.exSup3Place, 'city')
        self.exSup3Place.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.exSup3Place.setObjectName("exSup3Place")
        self.gridLayout_4.addWidget(self.exSup3Place, 11, 5, 1, 1)
        self.label_74 = QtWidgets.QLabel(self.page_4)
        self.label_74.setText("")
        self.label_74.setObjectName("label_74")
        self.gridLayout_4.addWidget(self.label_74, 14, 13, 1, 2)
        self.label_72 = QtWidgets.QLabel(self.page_4)
        self.label_72.setText("")
        self.label_72.setObjectName("label_72")
        self.gridLayout_4.addWidget(self.label_72, 14, 2, 1, 1)
        self.label_63 = QtWidgets.QLabel(self.page_4)
        self.label_63.setText("")
        self.label_63.setObjectName("label_63")
        self.gridLayout_4.addWidget(self.label_63, 2, 3, 1, 1)
        #布点供应商4所在地
        self.exSup4Place = QtWidgets.QComboBox(self.page_4)
        #更新定点地点选项
        self.exSup4Place.activated.connect(lambda:self.select_sup_option(self.exSup4Place, 'city'))

        #调用函数
        self.combo_option(self.exSup4Place, 'city')

        self.exSup4Place.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.exSup4Place.setObjectName("exSup4Place")
        self.gridLayout_4.addWidget(self.exSup4Place, 12, 5, 1, 1)
        #外协件信息录入，零件下拉框
        self.exPart = QComboBox(self.page_4)
        #当选项被单击时获取选项文本，采用getText函数
        #self.exPart.activated.connect(self.getText)
        #调用函数
        self.combo_option(self.exPart, 'parts')

        #筛选布点供应商选项
        self.exPart.activated.connect(lambda:self.exSup_option_byTech(self.exSup1))

        self.exPart.setMinimumSize(QtCore.QSize(100, 0))
        self.exPart.setMaximumSize(QtCore.QSize(130, 16777215))
        self.exPart.setObjectName("exPart")
        self.gridLayout_4.addWidget(self.exPart, 5, 6, 1, 2)
        #外协加信息录入，注塑发泡下拉框
        self.exInjection = QComboBox(self.page_4)
        #当选项被单击时获取选项文本，采用getText函数
        #self.exInjection.activated.connect(self.getText)
        #调用函数
        self.combo_option(self.exInjection, 'injection')

        #筛选布点供应商选项

        self.exInjection.activated.connect(lambda:self.exSup_option_byTech(self.exSup1))
        self.exInjection.setMaximumSize(QtCore.QSize(130, 16777215))
        self.exInjection.setObjectName("exInjection")
        self.gridLayout_4.addWidget(self.exInjection, 6, 6, 1, 2)
        self.label_58 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_58.setFont(font)
        self.label_58.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_58.setObjectName("label_58")
        self.gridLayout_4.addWidget(self.label_58, 12, 7, 1, 1)
        self.label_52 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_52.setFont(font)
        self.label_52.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_52.setObjectName("label_52")
        self.gridLayout_4.addWidget(self.label_52, 10, 7, 1, 1)
        self.exSeleSupTaMark = QtWidgets.QLineEdit(self.page_4)
        self.exSeleSupTaMark.setMaximumSize(QtCore.QSize(60, 16777215))
        self.exSeleSupTaMark.setObjectName("exSeleSupTaMark")
        self.gridLayout_4.addWidget(self.exSeleSupTaMark, 13, 8, 1, 1)
        self.label_55 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_55.setFont(font)
        self.label_55.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_55.setObjectName("label_55")
        self.gridLayout_4.addWidget(self.label_55, 11, 7, 1, 1)
        self.exSeleSupPlace = QtWidgets.QComboBox(self.page_4)
        #添加空选项用于初始化
        self.exSeleSupPlace.addItem('')
        self.exSeleSupPlace.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.exSeleSupPlace.setObjectName("exSeleSupPlace")
        self.gridLayout_4.addWidget(self.exSeleSupPlace, 13, 5, 1, 1)
        self.exSup3TaMark = QtWidgets.QLineEdit(self.page_4)
        self.exSup3TaMark.setMaximumSize(QtCore.QSize(60, 16777215))
        self.exSup3TaMark.setObjectName("exSup3TaMark")
        self.gridLayout_4.addWidget(self.exSup3TaMark, 11, 8, 1, 1)
        self.label_40 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_40.setFont(font)
        self.label_40.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_40.setObjectName("label_40")
        self.label_40.setMinimumSize(QtCore.QSize(90, 0))
        self.gridLayout_4.addWidget(self.label_40, 4, 0, 1, 1)
        self.exBackUp = QtWidgets.QPushButton(self.page_4)
        self.exBackUp.setObjectName("exBackUp")
        self.exBackUp.setFont(self.font)

        self.exBackUp.setMinimumSize(QtCore.QSize(90, 20))
        #设置函数------------------------------------------------------------
        self.exBackUp.clicked.connect(self.back_up_db)

        self.gridLayout_4.addWidget(self.exBackUp, 14, 12, 1, 1)
        self.label_62 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_62.setFont(font)
        self.label_62.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_62.setObjectName("label_62")
        self.gridLayout_4.addWidget(self.label_62, 13, 7, 1, 1)
        self.label_50 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_50.setFont(font)
        self.label_50.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_50.setObjectName("label_50")
        self.gridLayout_4.addWidget(self.label_50, 9, 7, 1, 1)
        self.exSup1TaMark = QtWidgets.QLineEdit(self.page_4)
        self.exSup1TaMark.setMaximumSize(QtCore.QSize(60, 16777215))
        self.exSup1TaMark.setObjectName("exSup1TaMark")
        self.gridLayout_4.addWidget(self.exSup1TaMark, 9, 8, 1, 1)
        self.exSup2TaMark = QtWidgets.QLineEdit(self.page_4)
        self.exSup2TaMark.setMaximumSize(QtCore.QSize(60, 16777215))
        self.exSup2TaMark.setObjectName("exSup2TaMark")
        self.gridLayout_4.addWidget(self.exSup2TaMark, 10, 8, 1, 1)
        self.exSup4TaMark = QtWidgets.QLineEdit(self.page_4)
        self.exSup4TaMark.setMaximumSize(QtCore.QSize(60, 16777215))
        self.exSup4TaMark.setObjectName("exSup4TaMark")
        self.gridLayout_4.addWidget(self.exSup4TaMark, 12, 8, 1, 1)
        self.label_77 = QtWidgets.QLabel(self.page_4)
        self.label_77.setText("")
        self.label_77.setObjectName("label_77")
        self.gridLayout_4.addWidget(self.label_77, 9, 13, 1, 2)
        self.label_38 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_38.setFont(font)
        self.label_38.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_38.setObjectName("label_38")
        self.label_38.setMinimumSize(QtCore.QSize(90, 0))
        self.gridLayout_4.addWidget(self.label_38, 3, 3, 1, 1)
        #布点供应商1
        self.exSup1 = QtWidgets.QComboBox(self.page_4)
        #更新定点选项
        self.exSup1.activated.connect(lambda:self.select_sup_option(self.exSup1, 'sup'))

        #self.combo_option(self.exSup1, 'abbreviation')
        self.exSup1.setMinimumSize(QtCore.QSize(100, 0))
        self.exSup1.setObjectName("exSup1")
        self.gridLayout_4.addWidget(self.exSup1, 9, 2, 1, 2)
        #外协件信息录入，装饰类下拉框
        self.exDecorate = QComboBox(self.page_4)
         #当选项被单击时获取选项文本，采用getText函数
        #self.exDecorate.activated.connect(self.getText)
        #调用函数
        self.combo_option(self.exDecorate, 'decoration')

        #筛选布点供应商选项
        self.exDecorate.activated.connect(lambda:self.exSup_option_byTech(self.exSup1))



        self.exDecorate.setMinimumSize(QtCore.QSize(0, 0))
        self.exDecorate.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exDecorate.setObjectName("exDecorate")
        self.gridLayout_4.addWidget(self.exDecorate, 6, 9, 1, 6)
        self.exProNumber = QtWidgets.QPushButton(self.page_4)
        self.exProNumber.setObjectName("exProNumber")
        self.exProNumber.setFont(self.font)
        self.gridLayout_4.addWidget(self.exProNumber, 2, 8, 1, 1)
        #外协件信息录入，部门下拉框
        self.exDepart = QtWidgets.QComboBox(self.page_4)
       
        self.exDepart.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exDepart.setObjectName("exDepart")
        self.exDepart.setMinimumSize(QtCore.QSize(167, 0))
        self.gridLayout_4.addWidget(self.exDepart, 3, 9, 1, 6)
         #调用函数
        self.combo_option(self.exDepart, 'depName')
        self.exInputDb = QtWidgets.QPushButton(self.page_4)
        self.exInputDb.setObjectName("exInputDb")
        self.exInputDb.setFont(self.font)
        #连接函数
        self.exInputDb.clicked.connect(self.ex_ta_inputDb)
        self.gridLayout_4.addWidget(self.exInputDb, 14, 4, 1, 1)
        self.line = QtWidgets.QFrame(self.page_4)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.gridLayout_4.addWidget(self.line, 8, 0, 1, 15)
        self.label_45 = QtWidgets.QLabel(self.page_4)
        self.label_45.setMinimumSize(QtCore.QSize(80, 22))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_45.setFont(font)
        self.label_45.setAlignment(QtCore.Qt.AlignLeft|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_45.setObjectName("label_45")
        self.gridLayout_4.addWidget(self.label_45, 6, 8, 1, 1)
        self.label_49 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_49.setFont(font)
        self.label_49.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_49.setObjectName("label_49")
        self.gridLayout_4.addWidget(self.label_49, 9, 4, 1, 1)
        self.label_66 = QtWidgets.QLabel(self.page_4)
        self.label_66.setText("")
        self.label_66.setObjectName("label_66")
        self.gridLayout_4.addWidget(self.label_66, 2, 14, 1, 1)
        self.label_73 = QtWidgets.QLabel(self.page_4)
        self.label_73.setText("")
        self.label_73.setObjectName("label_73")
        self.gridLayout_4.addWidget(self.label_73, 14, 5, 1, 1)
        self.label_64 = QtWidgets.QLabel(self.page_4)
        self.label_64.setText("")
        self.label_64.setObjectName("label_64")
        self.gridLayout_4.addWidget(self.label_64, 2, 5, 1, 1)
        #外协件信息录入OEM选框
        self.exOEM = QtWidgets.QComboBox(self.page_4)
        self.exOEM.setObjectName("exOEM")

        
         #调用函数
        self.combo_option(self.exOEM, 'oem')
        self.gridLayout_4.addWidget(self.exOEM, 3, 1, 1, 1)
        self.exSupDevQuaReview = QtWidgets.QPushButton(self.page_4)
        self.exSupDevQuaReview.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exSupDevQuaReview.setObjectName("exSupDevQuaReview")
        self.exSupDevQuaReview.setFont(self.font)
        #连接函数
        self.exSupDevQuaReview.clicked.connect(self.ex_sup_qua_review)


        self.gridLayout_4.addWidget(self.exSupDevQuaReview, 13, 9, 1, 5)
        self.exClear = QtWidgets.QPushButton(self.page_4)
        self.exClear.setObjectName("exClear")
        self.exClear.setFont(self.font)
        #连接清空函数
        self.exClear.clicked.connect(self.clear_all)
        self.gridLayout_4.addWidget(self.exClear, 14, 8, 1, 1)
        #外协件信息录入页，导入excel
        self.exAttach = QtWidgets.QPushButton(self.page_4)
        self.exAttach.setObjectName("exAttach")
        #连接函数
        self.exAttach.clicked.connect(self.ex_from_excel)

        self.exAttach.setFont(self.font)
        self.gridLayout_4.addWidget(self.exAttach, 2, 4, 1, 1)
        self.label_54 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_54.setFont(font)
        self.label_54.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_54.setObjectName("label_54")
        self.gridLayout_4.addWidget(self.label_54, 11, 4, 1, 1)
        self.label_51 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_51.setFont(font)
        self.label_51.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_51.setObjectName("label_51")
        self.gridLayout_4.addWidget(self.label_51, 10, 4, 1, 1)
        self.label_39 = QtWidgets.QLabel(self.page_4)
        self.label_39.setMinimumSize(QtCore.QSize(80, 22))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_39.setFont(font)
        #self.label_39.setAlignment(QtCore.Qt.AlignLeft|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_39.setObjectName("label_39")
        #self.label_39.setMaximumSize(QtCore.QSize(100, 0))
        self.gridLayout_4.addWidget(self.label_39, 3, 8, 1, 1)
        self.exProName = QtWidgets.QLineEdit(self.page_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.exProName.sizePolicy().hasHeightForWidth())
        self.exProName.setSizePolicy(sizePolicy)
        self.exProName.setMinimumSize(QtCore.QSize(150, 0))
        self.exProName.setMaximumSize(QtCore.QSize(176, 16777215))
        self.exProName.setObjectName("exProName")
        self.gridLayout_4.addWidget(self.exProName, 3, 4, 1, 2)
        self.label_57 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_57.setFont(font)
        self.label_57.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_57.setObjectName("label_57")
        self.gridLayout_4.addWidget(self.label_57, 12, 4, 1, 1)
        self.exTechReview = QtWidgets.QPushButton(self.page_4)
        self.exTechReview.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exTechReview.setObjectName("exTechReview")
        #连接界面跳转函数
        self.exTechReview.clicked.connect(self.ex_tech_review)

        self.exTechReview.setFont(self.font)
        self.gridLayout_4.addWidget(self.exTechReview, 11, 9, 1, 5)
        self.line_2 = QtWidgets.QFrame(self.page_4)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.gridLayout_4.addWidget(self.line_2, 15, 0, 1, 15)
        self.label_65 = QtWidgets.QLabel(self.page_4)
        self.label_65.setText("")
        self.label_65.setObjectName("label_65")
        self.gridLayout_4.addWidget(self.label_65, 2, 13, 1, 1)
        self.label_61 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_61.setFont(font)
        self.label_61.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_61.setObjectName("label_61")
        self.gridLayout_4.addWidget(self.label_61, 13, 4, 1, 1)


        self.exTa = QtWidgets.QLineEdit(self.page_4)



        #连接选项函数
        #self.combo_option(self.exTabox, 'TAer')

        self.exTa.setObjectName("exTa")
        #self.exTabox = QtWidgets.QComboBox(self.page_4)
        #self.exTabox.setObjectName("exTabox")
        #self.exTabox.addItem("")
        #self.combo_option(self.exTabox, 'TAer')

        self.gridLayout_4.addWidget(self.exTa, 6, 2, 1, 1)
        #self.horizontalLayout_2.addLayout(self.gridLayout_4)
        self.stackedWidget.addWidget(self.page_4)

       
        self.exInputOption = QtWidgets.QComboBox(self.page_4)
        self.exInputOption.setObjectName("exInputOption")
        #连接赋值函数
        self.exInputOption.activated.connect(self.ex_excel_to_ui)

        #self.exInputOption.addItem('try')
        self.gridLayout_4.addWidget(self.exInputOption, 4, 1, 1, 14)
        self.label_48 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_48.setFont(font)
        self.label_48.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_48.setObjectName("label_48")
        self.gridLayout_4.addWidget(self.label_48, 7, 0, 1, 2)
        self.label_47 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_47.setFont(font)
        self.label_47.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_47.setObjectName("label_47")
        self.gridLayout_4.addWidget(self.label_47, 5, 0, 1, 2)
        self.label_43 = QtWidgets.QLabel(self.page_4)
        self.label_43.setMinimumSize(QtCore.QSize(96, 22))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_43.setFont(font)
        self.label_43.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_43.setObjectName("label_43")
        self.gridLayout_4.addWidget(self.label_43, 5, 4, 1, 2)
        self.label_25 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_25.setFont(font)
        self.label_25.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_25.setObjectName("label_25")
        self.gridLayout_4.addWidget(self.label_25, 9, 0, 1, 2)
        self.label_53 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_53.setFont(font)
        self.label_53.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_53.setObjectName("label_53")
        self.gridLayout_4.addWidget(self.label_53, 10, 0, 1, 2)
        self.label_56 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_56.setFont(font)
        self.label_56.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_56.setObjectName("label_56")
        self.gridLayout_4.addWidget(self.label_56, 11, 0, 1, 2)
        self.label_60 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_60.setFont(font)
        self.label_60.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_60.setObjectName("label_60")
        self.gridLayout_4.addWidget(self.label_60, 13, 0, 1, 2)
        self.label_59 = QtWidgets.QLabel(self.page_4)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_59.setFont(font)
        self.label_59.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_59.setObjectName("label_59")
        self.gridLayout_4.addWidget(self.label_59, 12, 0, 1, 2)
        self.label_26 = QtWidgets.QLabel(self.page_4)
        self.label_26.setMaximumSize(QtCore.QSize(16777215, 30))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_26.setFont(font)
        self.label_26.setObjectName("label_26")
        self.gridLayout_4.addWidget(self.label_26, 2, 0, 1, 4)
        self.exTAtime = QtWidgets.QLineEdit(self.page_4)
        self.exTAtime.setObjectName("exTAtime")
        self.exTAtime.setMinimumSize(QtCore.QSize(96, 20))
        self.gridLayout_4.addWidget(self.exTAtime, 5, 2, 1, 1)
        self.exMouldTime = QtWidgets.QLineEdit(self.page_4)
        self.exMouldTime.setMinimumSize(QtCore.QSize(96, 20))
        self.exMouldTime.setObjectName("exMouldTime")
        self.gridLayout_4.addWidget(self.exMouldTime, 7, 2, 1, 1)
        self.label_85 = QtWidgets.QLabel(self.page_4)
        self.label_85.setMaximumSize(QtCore.QSize(16777215, 30))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_85.setFont(font)
        self.label_85.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_85.setObjectName("label_85")
        self.gridLayout_4.addWidget(self.label_85, 6, 0, 1, 2)
        #外协件信息录入页，金属成型下拉框
        self.exMetal = QComboBox(self.page_4)
         #当选项被单击时获取选项文本，采用getText函数
        #self.exMetal.activated.connect(self.getText)
        #调用函数
        self.combo_option(self.exMetal, 'metalForming')

        #筛选布点供应商选项
        self.exMetal.activated.connect(lambda:self.exSup_option_byTech(self.exSup1))


        self.exMetal.setMaximumSize(QtCore.QSize(130, 16777215))
        self.exMetal.setObjectName("exMetal")
        self.gridLayout_4.addWidget(self.exMetal, 7, 6, 1, 2)
        self.horizontalLayout_2.addLayout(self.gridLayout_4)
        self.stackedWidget.addWidget(self.page_4)
        #2019.7.27----------更新添加导出至excel按钮
        self.exOutputExcel = QtWidgets.QPushButton(self.page_4)

        #连接导出函数
        self.exOutputExcel.clicked.connect(self.ex_export_excel)

        self.exOutputExcel.setFont(self.font)
        self.exOutputExcel.setObjectName("exOutputExcel")
        self.gridLayout_4.addWidget(self.exOutputExcel, 14, 6, 1, 1)
        #self.horizontalLayout_2.addLayout(self.gridLayout_4)
        self.stackedWidget.addWidget(self.page_4)
        #第五页，待开发
        self.page_5 = QtWidgets.QWidget()
        self.page_5.setObjectName("page_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.page_5)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.label_27 = QtWidgets.QLabel(self.page_5)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_27.setFont(font)
        self.label_27.setAlignment(QtCore.Qt.AlignCenter)
        self.label_27.setObjectName("label_27")
        self.verticalLayout_5.addWidget(self.label_27)
        self.label_28 = QtWidgets.QLabel(self.page_5)
        self.label_28.setText("")
        self.label_28.setObjectName("label_28")
        self.verticalLayout_5.addWidget(self.label_28)
        self.stackedWidget.addWidget(self.page_5)
        self.gridLayout.addWidget(self.stackedWidget, 1, 1, 1, 1)
        #self.retranslateUi(Form)
        self.stackedWidget.setCurrentIndex(-1)

         #page6------------------------------------------------------------------------------2019.8.11
        self.page_6 = QtWidgets.QWidget()
        self.page_6.setObjectName("page_6")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.page_6)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.gridLayout_6 = QtWidgets.QGridLayout()
        self.gridLayout_6.setObjectName("gridLayout_6")

        #技术评估页，评审时间补充
        self.exTechReTime = QtWidgets.QLineEdit(self.page_6)
        self.exTechReTime.setObjectName("exTechReTime")
        self.gridLayout_6.addWidget(self.exTechReTime, 1, 12, 1, 1)

        self.label_107 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_107.setFont(font)
        self.label_107.setAlignment(QtCore.Qt.AlignCenter)
        self.label_107.setObjectName("label_107")
        self.gridLayout_6.addWidget(self.label_107, 1, 11, 1, 1)

        #计算分数按钮补充
        self.exTechCal = QtWidgets.QPushButton(self.page_6)

        #连接求和函数
        self.exTechCal.clicked.connect(self.cal_tech_mark)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(9)
        self.exTechCal.setFont(font)
        self.exTechCal.setObjectName("exTechCal")
        self.gridLayout_6.addWidget(self.exTechCal, 11, 7, 1, 1)

        self.label_88 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_88.setFont(font)
        self.label_88.setObjectName("label_88")
        self.gridLayout_6.addWidget(self.label_88, 2, 0, 1, 2)
        self.label_91 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_91.setFont(font)
        self.label_91.setObjectName("label_91")
        self.gridLayout_6.addWidget(self.label_91, 5, 1, 1, 1)
        self.label_80 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_80.setFont(font)
        self.label_80.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_80.setObjectName("label_80")
        self.gridLayout_6.addWidget(self.label_80, 0, 4, 1, 3)
        self.exTechCity = QtWidgets.QLineEdit(self.page_6)
        self.exTechCity.setObjectName("exTechCity")
        self.gridLayout_6.addWidget(self.exTechCity, 0, 10, 1, 1)
        self.label_81 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_81.setFont(font)
        self.label_81.setObjectName("label_81")
        self.gridLayout_6.addWidget(self.label_81, 0, 9, 1, 1)
        self.label_82 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_82.setFont(font)
        self.label_82.setObjectName("label_82")
        self.gridLayout_6.addWidget(self.label_82, 0, 11, 1, 1)
        self.label_83 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_83.setFont(font)
        self.label_83.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_83.setObjectName("label_83")
        self.gridLayout_6.addWidget(self.label_83, 1, 1, 1, 1)
        self.label_87 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_87.setFont(font)
        self.label_87.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_87.setObjectName("label_87")
        self.gridLayout_6.addWidget(self.label_87, 1, 4, 1, 3)
        self.exTechProvin = QtWidgets.QLineEdit(self.page_6)
        self.exTechProvin.setObjectName("exTechProvin")
        self.gridLayout_6.addWidget(self.exTechProvin, 0, 7, 1, 2)
        self.label_89 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_89.setFont(font)
        self.label_89.setObjectName("label_89")
        self.gridLayout_6.addWidget(self.label_89, 4, 1, 1, 1)
        self.line_4 = QtWidgets.QFrame(self.page_6)
        self.line_4.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.gridLayout_6.addWidget(self.line_4, 7, 0, 1, 13)
        self.label_94 = QtWidgets.QLabel(self.page_6)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_94.sizePolicy().hasHeightForWidth())
        self.label_94.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_94.setFont(font)
        self.label_94.setAlignment(QtCore.Qt.AlignCenter)
        self.label_94.setObjectName("label_94")
        self.gridLayout_6.addWidget(self.label_94, 8, 6, 1, 3)
        self.exTechClear = QtWidgets.QPushButton(self.page_6)

        self.exTechClear.setMaximumSize(QtCore.QSize(90, 30))

        #连接函数
        self.exTechClear.clicked.connect(self.clear_all)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exTechClear.setFont(font)
        self.exTechClear.setObjectName("exTechClear")
        self.gridLayout_6.addWidget(self.exTechClear, 12, 11, 1, 1)
        self.exTechTotalMark = QtWidgets.QLineEdit(self.page_6)
        self.exTechTotalMark.setObjectName("exTechTotalMark")
        self.gridLayout_6.addWidget(self.exTechTotalMark, 11, 5, 1, 2)
        self.label_93 = QtWidgets.QLabel(self.page_6)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_93.sizePolicy().hasHeightForWidth())
        self.label_93.setSizePolicy(sizePolicy)
        self.label_93.setMinimumSize(QtCore.QSize(0, 21))
        self.label_93.setMaximumSize(QtCore.QSize(16777215, 14444444))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_93.setFont(font)
        self.label_93.setAlignment(QtCore.Qt.AlignCenter)
        self.label_93.setObjectName("label_93")
        self.gridLayout_6.addWidget(self.label_93, 8, 2, 1, 3)
        self.tabWidget = QtWidgets.QTabWidget(self.page_6)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.textBrowser = QtWidgets.QTextBrowser(self.tab)
        self.textBrowser.setGeometry(QtCore.QRect(210, 60, 211, 81))
        self.textBrowser.setObjectName("textBrowser")
        self.label_97 = QtWidgets.QLabel(self.tab)
        self.label_97.setGeometry(QtCore.QRect(30, 90, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_97.setFont(font)
        self.label_97.setObjectName("label_97")
        self.exTechM1 = QtWidgets.QLineEdit(self.tab)
        self.exTechM1.setGeometry(QtCore.QRect(440, 60, 51, 81))
        self.exTechM1.setObjectName("exTechM1")
        self.exTechDesF1 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesF1.setGeometry(QtCore.QRect(510, 60, 211, 81))
        self.exTechDesF1.setObjectName("exTechDesF1")
        self.exTechDesA1 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesA1.setGeometry(QtCore.QRect(740, 60, 211, 81))
        self.exTechDesA1.setObjectName("exTechDesA1")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.tab)
        self.textBrowser_2.setGeometry(QtCore.QRect(210, 150, 211, 61))
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.exTechM2 = QtWidgets.QLineEdit(self.tab)
        self.exTechM2.setGeometry(QtCore.QRect(440, 150, 51, 61))
        self.exTechM2.setObjectName("exTechM2")
        self.exTechDesF2 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesF2.setGeometry(QtCore.QRect(510, 150, 211, 61))
        self.exTechDesF2.setObjectName("exTechDesF2")
        self.exTechDesA2 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesA2.setGeometry(QtCore.QRect(740, 150, 211, 61))
        self.exTechDesA2.setObjectName("exTechDesA2")
        self.label_99 = QtWidgets.QLabel(self.tab)
        self.label_99.setGeometry(QtCore.QRect(20, 240, 191, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_99.setFont(font)
        self.label_99.setObjectName("label_99")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.tab)
        self.textBrowser_3.setGeometry(QtCore.QRect(210, 230, 211, 81))
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.exTechM3 = QtWidgets.QLineEdit(self.tab)
        self.exTechM3.setGeometry(QtCore.QRect(440, 230, 51, 81))
        self.exTechM3.setObjectName("exTechM3")
        self.exTechDesF3 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesF3.setGeometry(QtCore.QRect(510, 230, 211, 81))
        self.exTechDesF3.setObjectName("exTechDesF3")
        self.exTechDesA3 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesA3.setGeometry(QtCore.QRect(740, 230, 211, 81))
        self.exTechDesA3.setObjectName("exTechDesA3")
        self.label_100 = QtWidgets.QLabel(self.tab)
        self.label_100.setGeometry(QtCore.QRect(10, 350, 201, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_100.setFont(font)
        self.label_100.setObjectName("label_100")
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.tab)
        self.textBrowser_4.setGeometry(QtCore.QRect(210, 330, 211, 61))
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.exTechM4 = QtWidgets.QLineEdit(self.tab)
        self.exTechM4.setGeometry(QtCore.QRect(440, 330, 51, 61))
        self.exTechM4.setObjectName("exTechM4")
        self.exTechDesF4 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesF4.setGeometry(QtCore.QRect(510, 330, 211, 61))
        self.exTechDesF4.setObjectName("exTechDesF4")
        self.exTechDesA4 = QtWidgets.QTextEdit(self.tab)
        self.exTechDesA4.setGeometry(QtCore.QRect(740, 330, 211, 61))
        self.exTechDesA4.setObjectName("exTechDesA4")
        #self.label_98 = QtWidgets.QLabel(self.tab)
        #self.label_98.setGeometry(QtCore.QRect(60, 90, 41, 118))
        #self.label_98.setText("")
        #self.label_98.setObjectName("label_98")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.textBrowser_5 = QtWidgets.QTextBrowser(self.tab_2)
        self.textBrowser_5.setGeometry(QtCore.QRect(140, 10, 241, 111))
        self.textBrowser_5.setObjectName("textBrowser_5")
        self.exTechM5 = QtWidgets.QLineEdit(self.tab_2)
        self.exTechM5.setGeometry(QtCore.QRect(430, 10, 51, 111))
        self.exTechM5.setObjectName("exTechM5")
        self.exTechDesF5 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesF5.setGeometry(QtCore.QRect(530, 10, 211, 111))
        self.exTechDesF5.setObjectName("exTechDesF5")
        self.exTechDesA5 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesA5.setGeometry(QtCore.QRect(770, 10, 211, 111))
        self.exTechDesA5.setObjectName("exTechDesA5")
        self.textBrowser_6 = QtWidgets.QTextBrowser(self.tab_2)
        self.textBrowser_6.setGeometry(QtCore.QRect(140, 130, 241, 111))
        self.textBrowser_6.setObjectName("textBrowser_6")
        self.exTechM6 = QtWidgets.QLineEdit(self.tab_2)
        self.exTechM6.setGeometry(QtCore.QRect(430, 130, 51, 111))
        self.exTechM6.setObjectName("exTechM6")
        self.exTechDesF6 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesF6.setGeometry(QtCore.QRect(530, 130, 211, 111))
        self.exTechDesF6.setObjectName("exTechDesF6")
        self.exTechDesA6 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesA6.setGeometry(QtCore.QRect(770, 130, 211, 111))
        self.exTechDesA6.setObjectName("exTechDesA6")
        self.textBrowser_7 = QtWidgets.QTextBrowser(self.tab_2)
        self.textBrowser_7.setGeometry(QtCore.QRect(140, 250, 241, 81))
        self.textBrowser_7.setObjectName("textBrowser_7")
        self.exTechM7 = QtWidgets.QLineEdit(self.tab_2)
        self.exTechM7.setGeometry(QtCore.QRect(430, 260, 51, 71))
        self.exTechM7.setObjectName("exTechM7")
        self.exTechDesF7 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesF7.setGeometry(QtCore.QRect(530, 260, 211, 71))
        self.exTechDesF7.setObjectName("exTechDesF7")
        self.exTechDesA7 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesA7.setGeometry(QtCore.QRect(770, 260, 211, 71))
        self.exTechDesA7.setObjectName("exTechDesA7")
        self.textBrowser_8 = QtWidgets.QTextBrowser(self.tab_2)
        self.textBrowser_8.setGeometry(QtCore.QRect(140, 340, 241, 81))
        self.textBrowser_8.setObjectName("textBrowser_8")
        self.exTechM8 = QtWidgets.QLineEdit(self.tab_2)
        self.exTechM8.setGeometry(QtCore.QRect(430, 350, 51, 71))
        self.exTechM8.setObjectName("exTechM8")
        self.exTechDesF8 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesF8.setGeometry(QtCore.QRect(530, 350, 211, 71))
        self.exTechDesF8.setObjectName("exTechDesF8")
        self.exTechDesA8 = QtWidgets.QTextEdit(self.tab_2)
        self.exTechDesA8.setGeometry(QtCore.QRect(770, 350, 211, 71))
        self.exTechDesA8.setObjectName("exTechDesA8")
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.textBrowser_9 = QtWidgets.QTextBrowser(self.tab_3)
        self.textBrowser_9.setGeometry(QtCore.QRect(190, 10, 271, 111))
        self.textBrowser_9.setObjectName("textBrowser_9")
        self.exTechM9 = QtWidgets.QLineEdit(self.tab_3)
        self.exTechM9.setGeometry(QtCore.QRect(470, 10, 51, 111))
        self.exTechM9.setObjectName("exTechM9")
        self.exTechDesF9 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesF9.setGeometry(QtCore.QRect(530, 10, 211, 111))
        self.exTechDesF9.setObjectName("exTechDesF9")
        self.exTechDesA9 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesA9.setGeometry(QtCore.QRect(750, 10, 211, 111))
        self.exTechDesA9.setObjectName("exTechDesA9")
        self.textBrowser_10 = QtWidgets.QTextBrowser(self.tab_3)
        self.textBrowser_10.setGeometry(QtCore.QRect(190, 130, 271, 81))
        self.textBrowser_10.setObjectName("textBrowser_10")
        self.exTechM10 = QtWidgets.QLineEdit(self.tab_3)
        self.exTechM10.setGeometry(QtCore.QRect(470, 130, 51, 81))
        self.exTechM10.setObjectName("exTechM10")
        self.exTechDesF10 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesF10.setGeometry(QtCore.QRect(530, 130, 211, 81))
        self.exTechDesF10.setObjectName("exTechDesF10")
        self.exTechDesA10 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesA10.setGeometry(QtCore.QRect(750, 130, 211, 81))
        self.exTechDesA10.setObjectName("exTechDesA10")
        self.textBrowser_11 = QtWidgets.QTextBrowser(self.tab_3)
        self.textBrowser_11.setGeometry(QtCore.QRect(190, 240, 271, 41))
        self.textBrowser_11.setObjectName("textBrowser_11")
        self.exTechM11 = QtWidgets.QLineEdit(self.tab_3)
        self.exTechM11.setGeometry(QtCore.QRect(470, 240, 51, 41))
        self.exTechM11.setObjectName("exTechM11")
        self.exTechDesF11 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesF11.setGeometry(QtCore.QRect(530, 240, 211, 41))
        self.exTechDesF11.setObjectName("exTechDesF11")
        self.exTechDesA11 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesA11.setGeometry(QtCore.QRect(750, 240, 211, 41))
        self.exTechDesA11.setObjectName("exTechDesA11")
        self.textBrowser_12 = QtWidgets.QTextBrowser(self.tab_3)
        self.textBrowser_12.setGeometry(QtCore.QRect(190, 300, 271, 61))
        self.textBrowser_12.setObjectName("textBrowser_12")
        self.exTechM12 = QtWidgets.QLineEdit(self.tab_3)
        self.exTechM12.setGeometry(QtCore.QRect(470, 300, 51, 51))
        self.exTechM12.setObjectName("exTechM12")
        self.exTechDesF12 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesF12.setGeometry(QtCore.QRect(530, 300, 211, 51))
        self.exTechDesF12.setObjectName("exTechDesF12")
        self.exTechDesA12 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesA12.setGeometry(QtCore.QRect(750, 300, 211, 51))
        self.exTechDesA12.setObjectName("exTechDesA12")
        self.textBrowser_13 = QtWidgets.QTextBrowser(self.tab_3)
        self.textBrowser_13.setGeometry(QtCore.QRect(190, 370, 271, 51))
        self.textBrowser_13.setObjectName("textBrowser_13")
        self.exTechM13 = QtWidgets.QLineEdit(self.tab_3)
        self.exTechM13.setGeometry(QtCore.QRect(470, 370, 51, 51))
        self.exTechM13.setObjectName("exTechM13")
        self.exTechDesF13 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesF13.setGeometry(QtCore.QRect(530, 370, 211, 51))
        self.exTechDesF13.setObjectName("exTechDesF13")
        self.exTechDesA13 = QtWidgets.QTextEdit(self.tab_3)
        self.exTechDesA13.setGeometry(QtCore.QRect(750, 370, 211, 51))
        self.exTechDesA13.setObjectName("exTechDesA13")
        self.label_109 = QtWidgets.QLabel(self.tab_3)
        self.label_109.setGeometry(QtCore.QRect(20, 380, 171, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_109.setFont(font)
        self.label_109.setObjectName("label_109")
        self.tabWidget.addTab(self.tab_3, "")
        self.gridLayout_6.addWidget(self.tabWidget, 9, 0, 2, 13)
        self.exTechImport = QtWidgets.QPushButton(self.page_6)

        #连接函数
        self.exTechImport.clicked.connect(self.teche_review_from_excel)

        self.exTechImport.setMaximumSize(QtCore.QSize(150, 30))

        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exTechImport.setFont(font)
        self.exTechImport.setObjectName("exTechImport")
        self.gridLayout_6.addWidget(self.exTechImport, 12, 3, 1, 3)
        self.exTechInputDb = QtWidgets.QPushButton(self.page_6)

        #连接函数
        self.exTechInputDb.clicked.connect(self.tech_review_to_db)

        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exTechInputDb.setFont(font)
        self.exTechInputDb.setObjectName("exTechInputDb")
        self.gridLayout_6.addWidget(self.exTechInputDb, 12, 8, 1, 2)
        self.label_110 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_110.setFont(font)
        self.label_110.setObjectName("label_110")
        self.gridLayout_6.addWidget(self.label_110, 11, 2, 1, 2)
        self.exTechBack = QtWidgets.QPushButton(self.page_6)

        self.exTechBack.setMaximumSize(QtCore.QSize(90, 30))

        #连接函数
        self.exTechBack.clicked.connect(self.back_to_external)

        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exTechBack.setFont(font)
        self.exTechBack.setObjectName("exTechBack")
        self.gridLayout_6.addWidget(self.exTechBack, 12, 12, 1, 1)
        self.exTechInjection = QtWidgets.QComboBox(self.page_6)
        self.exTechInjection.setObjectName("exTechInjection")

        #连接选项函数
        self.combo_option(self.exTechInjection, 'injection')

        self.gridLayout_6.addWidget(self.exTechInjection, 4, 2, 1, 4)
        self.exTechPart = QtWidgets.QComboBox(self.page_6)
        self.exTechPart.setObjectName("exTechPart")

        #连接选项函数
        self.combo_option(self.exTechPart, 'parts')

        self.gridLayout_6.addWidget(self.exTechPart, 2, 2, 1, 4)
        self.exTechWield = QtWidgets.QComboBox(self.page_6)
        self.exTechWield.setObjectName("exTechWield")

        #连接函数
        self.combo_option(self.exTechWield, 'welding')

        self.gridLayout_6.addWidget(self.exTechWield, 5, 10, 1, 1)
        self.exTechMetal = QtWidgets.QComboBox(self.page_6)
        self.exTechMetal.setObjectName("exTechMetal")

        #连接函数
        self.combo_option(self.exTechMetal, 'metalForming')

        self.gridLayout_6.addWidget(self.exTechMetal, 5, 2, 1, 4)
        self.exTechDecorate = QtWidgets.QComboBox(self.page_6)

        #连接函数
        self.combo_option(self.exTechDecorate, 'decoration')

        self.exTechDecorate.setObjectName("exTechDecorate")
        self.gridLayout_6.addWidget(self.exTechDecorate, 4, 10, 1, 1)
        self.label_95 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_95.setFont(font)
        self.label_95.setObjectName("label_95")
        self.gridLayout_6.addWidget(self.label_95, 8, 9, 1, 2)
        self.label_90 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_90.setFont(font)
        self.label_90.setObjectName("label_90")
        self.gridLayout_6.addWidget(self.label_90, 4, 8, 1, 2)
        self.exTechProName = QtWidgets.QLineEdit(self.page_6)
        self.exTechProName.setObjectName("exTechProName")
        self.gridLayout_6.addWidget(self.exTechProName, 1, 2, 1, 2)
        self.exTechSupName = QtWidgets.QLineEdit(self.page_6)
        self.exTechSupName.setObjectName("exTechSupName")
        self.gridLayout_6.addWidget(self.exTechSupName, 0, 2, 1, 2)
        self.label_92 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_92.setFont(font)
        self.label_92.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_92.setObjectName("label_92")
        self.gridLayout_6.addWidget(self.label_92, 5, 8, 1, 2)
        self.label_63 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(12)
        self.label_63.setFont(font)
        self.label_63.setObjectName("label_63")
        self.gridLayout_6.addWidget(self.label_63, 4, 12, 1, 1)
        self.exTechSupNum = QtWidgets.QComboBox(self.page_6)

        #连接函数，确定布点供应商名称
        self.exTechSupNum.activated.connect(self.contentReference_from_external_to_techReciew)
        self.exTechSupNum.setObjectName("exTechSupNum")
        self.exTechSupNum.addItem("")
        self.exTechSupNum.addItem("")
        self.exTechSupNum.addItem("")
        self.exTechSupNum.addItem("")
        self.exTechSupNum.addItem("")
        self.gridLayout_6.addWidget(self.exTechSupNum, 5, 12, 1, 1)
        self.exTechRevProduct = QtWidgets.QTextEdit(self.page_6)
        self.exTechRevProduct.setObjectName("exTechRevProduct")
        self.gridLayout_6.addWidget(self.exTechRevProduct, 1, 7, 2, 4)
        self.label_34 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_34.setFont(font)
        self.label_34.setObjectName("label_34")
        self.gridLayout_6.addWidget(self.label_34, 0, 1, 1, 1)
        self.label_96 = QtWidgets.QLabel(self.page_6)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_96.setFont(font)
        self.label_96.setObjectName("label_96")
        self.gridLayout_6.addWidget(self.label_96, 8, 11, 1, 1)
        self.verticalLayout_6.addLayout(self.gridLayout_6)
        self.stackedWidget.addWidget(self.page_6)

        #第七页-------------------------------------------------------------供应商质量评估界面
        self.page_7 = QtWidgets.QWidget()
        self.page_7.setObjectName("page_7")
        self.verticalLayout_7 = QtWidgets.QVBoxLayout(self.page_7)
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.gridLayout_7 = QtWidgets.QGridLayout()
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.exQuaWield = QtWidgets.QComboBox(self.page_7)

        #下拉选项
        self.combo_option(self.exQuaWield, 'welding')

        self.exQuaWield.setObjectName("exQuaWield")
        self.gridLayout_7.addWidget(self.exQuaWield, 3, 9, 1, 3)
        self.label_115 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_115.setFont(font)
        self.label_115.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_115.setObjectName("label_115")
        self.gridLayout_7.addWidget(self.label_115, 3, 8, 1, 1)
        self.label_112 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_112.setFont(font)
        self.label_112.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_112.setObjectName("label_112")
        self.gridLayout_7.addWidget(self.label_112, 3, 3, 1, 1)
        self.label_114 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_114.setFont(font)
        self.label_114.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_114.setObjectName("label_114")
        self.gridLayout_7.addWidget(self.label_114, 2, 8, 1, 1)
        self.exQuaReProduct = QtWidgets.QTextEdit(self.page_7)
        self.exQuaReProduct.setObjectName("exQuaReProduct")
        self.gridLayout_7.addWidget(self.exQuaReProduct, 2, 1, 3, 2)
        self.label_120 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_120.setFont(font)
        self.label_120.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_120.setObjectName("label_120")
        self.gridLayout_7.addWidget(self.label_120, 1, 3, 1, 1)
        self.label_118 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_118.setFont(font)
        self.label_118.setObjectName("label_118")
        self.gridLayout_7.addWidget(self.label_118, 0, 9, 1, 1)
        self.label_116 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_116.setFont(font)
        self.label_116.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_116.setObjectName("label_116")
        self.gridLayout_7.addWidget(self.label_116, 0, 3, 1, 1)
        self.exQuaCity = QtWidgets.QLineEdit(self.page_7)
        self.exQuaCity.setObjectName("exQuaCity")
        self.gridLayout_7.addWidget(self.exQuaCity, 0, 8, 1, 1)
        self.label_35 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_35.setFont(font)
        self.label_35.setObjectName("label_35")
        self.gridLayout_7.addWidget(self.label_35, 0, 0, 1, 2)
        self.exQuaProName = QtWidgets.QLineEdit(self.page_7)
        self.exQuaProName.setObjectName("exQuaProName")
        self.gridLayout_7.addWidget(self.exQuaProName, 1, 2, 1, 1)
        self.label_117 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_117.setFont(font)
        self.label_117.setObjectName("label_117")
        self.gridLayout_7.addWidget(self.label_117, 0, 6, 1, 1)
        self.exQuaProNum = QtWidgets.QLineEdit(self.page_7)
        self.exQuaProNum.setObjectName("exQuaProNum")
        self.gridLayout_7.addWidget(self.exQuaProNum, 1, 4, 1, 4)
        self.label_119 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_119.setFont(font)
        self.label_119.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_119.setObjectName("label_119")
        self.gridLayout_7.addWidget(self.label_119, 1, 0, 1, 2)
        self.exQuaSupName = QtWidgets.QLineEdit(self.page_7)

        

        self.gridLayout_7.addWidget(self.exQuaSupName, 0, 2, 1, 1)
        self.label_84 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_84.setFont(font)
        self.label_84.setObjectName("label_84")
        self.gridLayout_7.addWidget(self.label_84, 2, 0, 1, 1)
        self.exQuaProvin = QtWidgets.QLineEdit(self.page_7)
        self.exQuaProvin.setObjectName("exQuaProvin")
        self.gridLayout_7.addWidget(self.exQuaProvin, 0, 4, 1, 1)
        self.tabWidget_2 = QtWidgets.QTabWidget(self.page_7)
        self.tabWidget_2.setObjectName("tabWidget_2")
        self.tab_4 = QtWidgets.QWidget()
        self.tab_4.setObjectName("tab_4")
        self.label_69 = QtWidgets.QLabel(self.tab_4)
        self.label_69.setGeometry(QtCore.QRect(110, 60, 91, 21))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_69.setFont(font)
        self.label_69.setObjectName("label_69")
        self.textBrowser_14 = QtWidgets.QTextBrowser(self.tab_4)
        self.textBrowser_14.setGeometry(QtCore.QRect(260, 30, 391, 81))
        self.textBrowser_14.setObjectName("textBrowser_14")
        self.exQuaM1 = QtWidgets.QLineEdit(self.tab_4)
        self.exQuaM1.setGeometry(QtCore.QRect(740, 30, 61, 81))
        self.exQuaM1.setObjectName("exQuaM1")
        self.textBrowser_15 = QtWidgets.QTextBrowser(self.tab_4)
        self.textBrowser_15.setGeometry(QtCore.QRect(260, 150, 391, 91))
        self.textBrowser_15.setObjectName("textBrowser_15")
        self.exQuaM2 = QtWidgets.QLineEdit(self.tab_4)
        self.exQuaM2.setGeometry(QtCore.QRect(740, 150, 61, 81))
        self.exQuaM2.setObjectName("exQuaM2")
        self.textBrowser_16 = QtWidgets.QTextBrowser(self.tab_4)
        self.textBrowser_16.setGeometry(QtCore.QRect(260, 270, 391, 81))
        self.textBrowser_16.setObjectName("textBrowser_16")
        self.exQuaM3 = QtWidgets.QLineEdit(self.tab_4)
        self.exQuaM3.setGeometry(QtCore.QRect(740, 270, 61, 81))
        self.exQuaM3.setObjectName("exQuaM3")
        self.label_121 = QtWidgets.QLabel(self.tab_4)
        self.label_121.setGeometry(QtCore.QRect(530, 450, 163, 21))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_121.setFont(font)
        self.label_121.setObjectName("label_121")
        self.exQuaTotal = QtWidgets.QLineEdit(self.tab_4)
        self.exQuaTotal.setGeometry(QtCore.QRect(740, 450, 61, 21))
        self.exQuaTotal.setObjectName("exQuaTotal")
        self.exSupCal = QtWidgets.QPushButton(self.tab_4)

        #连接计算函数
        self.exSupCal.clicked.connect(self.cal_qua_mark)

        self.exSupCal.setGeometry(QtCore.QRect(820, 450, 51, 23))
        self.exSupCal.setObjectName("exSupCal")
        self.tabWidget_2.addTab(self.tab_4, "")
        self.tab_5 = QtWidgets.QWidget()
        self.tab_5.setObjectName("tab_5")
        self.label_70 = QtWidgets.QLabel(self.tab_5)
        self.label_70.setGeometry(QtCore.QRect(80, 50, 111, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_70.setFont(font)
        self.label_70.setObjectName("label_70")
        self.textBrowser_17 = QtWidgets.QTextBrowser(self.tab_5)
        self.textBrowser_17.setGeometry(QtCore.QRect(270, 40, 361, 41))
        self.textBrowser_17.setObjectName("textBrowser_17")
        self.exQuaM4 = QtWidgets.QLineEdit(self.tab_5)
        self.exQuaM4.setGeometry(QtCore.QRect(740, 40, 61, 41))
        self.exQuaM4.setObjectName("exQuaM4")
        self.textBrowser_18 = QtWidgets.QTextBrowser(self.tab_5)
        self.textBrowser_18.setGeometry(QtCore.QRect(270, 130, 361, 41))
        self.textBrowser_18.setObjectName("textBrowser_18")
        self.exQuaM5 = QtWidgets.QLineEdit(self.tab_5)
        self.exQuaM5.setGeometry(QtCore.QRect(740, 130, 61, 41))
        self.exQuaM5.setObjectName("exQuaM5")
        self.textBrowser_19 = QtWidgets.QTextBrowser(self.tab_5)
        self.textBrowser_19.setGeometry(QtCore.QRect(270, 220, 361, 41))
        self.textBrowser_19.setObjectName("textBrowser_19")
        self.exQuaM6 = QtWidgets.QLineEdit(self.tab_5)
        self.exQuaM6.setGeometry(QtCore.QRect(740, 220, 61, 41))
        self.exQuaM6.setObjectName("exQuaM6")
        self.textBrowser_20 = QtWidgets.QTextBrowser(self.tab_5)
        self.textBrowser_20.setGeometry(QtCore.QRect(270, 320, 361, 81))
        self.textBrowser_20.setObjectName("textBrowser_20")
        self.exQuaM7 = QtWidgets.QLineEdit(self.tab_5)
        self.exQuaM7.setGeometry(QtCore.QRect(740, 320, 61, 81))
        self.exQuaM7.setObjectName("exQuaM7")
        self.tabWidget_2.addTab(self.tab_5, "")
        self.tab_6 = QtWidgets.QWidget()
        self.tab_6.setObjectName("tab_6")
        self.textBrowser_21 = QtWidgets.QTextBrowser(self.tab_6)
        self.textBrowser_21.setGeometry(QtCore.QRect(300, 70, 291, 131))
        self.textBrowser_21.setObjectName("textBrowser_21")
        self.label_78 = QtWidgets.QLabel(self.tab_6)
        self.label_78.setGeometry(QtCore.QRect(90, 80, 151, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_78.setFont(font)
        self.label_78.setObjectName("label_78")
        self.exQuaM8 = QtWidgets.QLineEdit(self.tab_6)
        self.exQuaM8.setGeometry(QtCore.QRect(740, 70, 61, 131))
        self.exQuaM8.setObjectName("exQuaM8")
        self.tabWidget_2.addTab(self.tab_6, "")
        self.gridLayout_7.addWidget(self.tabWidget_2, 7, 0, 2, 12)
        self.label_113 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_113.setFont(font)
        self.label_113.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_113.setObjectName("label_113")
        self.gridLayout_7.addWidget(self.label_113, 4, 3, 1, 1)
        self.line_3 = QtWidgets.QFrame(self.page_7)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.gridLayout_7.addWidget(self.line_3, 5, 0, 1, 12)
        self.label_41 = QtWidgets.QLabel(self.page_7)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_41.sizePolicy().hasHeightForWidth())
        self.label_41.setSizePolicy(sizePolicy)
        self.label_41.setMinimumSize(QtCore.QSize(0, 21))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_41.setFont(font)
        self.label_41.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_41.setObjectName("label_41")
        self.gridLayout_7.addWidget(self.label_41, 6, 2, 1, 1)
        self.label_111 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_111.setFont(font)
        self.label_111.setObjectName("label_111")
        self.gridLayout_7.addWidget(self.label_111, 2, 3, 1, 1)
        self.exQuaPart = QtWidgets.QComboBox(self.page_7)

        #下拉选项函数
        self.combo_option(self.exQuaPart, 'parts')

        self.exQuaPart.setObjectName("exQuaPart")
        self.gridLayout_7.addWidget(self.exQuaPart, 2, 4, 1, 4)
        self.exQuaDecorat = QtWidgets.QComboBox(self.page_7)

        #下拉选项
        self.combo_option(self.exQuaDecorat, 'decoration')

        self.exQuaDecorat.setObjectName("exQuaDecorat")
        self.gridLayout_7.addWidget(self.exQuaDecorat, 2, 9, 1, 3)
        self.exQuaInjection = QtWidgets.QComboBox(self.page_7)

        #选项函数
        self.combo_option(self.exQuaInjection, 'injection')

        self.exQuaInjection.setObjectName("exQuaInjection")
        self.gridLayout_7.addWidget(self.exQuaInjection, 3, 4, 1, 4)
        self.exQuaMetal = QtWidgets.QComboBox(self.page_7)

        #选项函数
        self.combo_option(self.exQuaMetal, 'metalForming')

        self.exQuaMetal.setObjectName("exQuaMetal")
        self.gridLayout_7.addWidget(self.exQuaMetal, 4, 4, 1, 4)

        self.exQuaImport = QtWidgets.QPushButton(self.page_7)

        #连接导入函数
        self.exQuaImport.clicked.connect(self.exQua_import_excel)
        self.exQuaImport.setMaximumSize(QtCore.QSize(150, 30))

        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exQuaImport.setFont(font)
        self.exQuaImport.setObjectName("exQuaImport")
        self.gridLayout_7.addWidget(self.exQuaImport, 9, 2, 1, 1)
        self.exQuaInputDb = QtWidgets.QPushButton(self.page_7)

        #连接导入函数
        self.exQuaInputDb.clicked.connect(self.qua_to_db)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exQuaInputDb.setFont(font)
        self.exQuaInputDb.setObjectName("exQuaInputDb")
        self.gridLayout_7.addWidget(self.exQuaInputDb, 9, 3, 1, 1)
        self.exQuaClear = QtWidgets.QPushButton(self.page_7)

        #连接清空函数
        self.exQuaClear.clicked.connect(self.clear_all)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exQuaClear.setFont(font)
        self.exQuaClear.setObjectName("exQuaClear")
        self.gridLayout_7.addWidget(self.exQuaClear, 9, 5, 1, 3)
        self.exTechBack_2 = QtWidgets.QPushButton(self.page_7)

        #连接返回函数
        self.exTechBack_2.clicked.connect(self.qua_back_tech)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.exTechBack_2.setFont(font)
        self.exTechBack_2.setObjectName("exTechBack_2")
        self.gridLayout_7.addWidget(self.exTechBack_2, 9, 9, 1, 2)
        self.label_67 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_67.setFont(font)
        self.label_67.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_67.setObjectName("label_67")
        self.gridLayout_7.addWidget(self.label_67, 6, 3, 1, 1)
        self.label_139 = QtWidgets.QLabel(self.page_7)
        self.label_139.setMinimumSize(QtCore.QSize(87, 21))
        self.label_139.setText("")
        self.label_139.setObjectName("label_139")
        self.gridLayout_7.addWidget(self.label_139, 4, 11, 1, 1)
        self.label_68 = QtWidgets.QLabel(self.page_7)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_68.setFont(font)
        self.label_68.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_68.setObjectName("label_68")
        self.gridLayout_7.addWidget(self.label_68, 6, 7, 1, 2)
        self.verticalLayout_7.addLayout(self.gridLayout_7)
        self.stackedWidget.addWidget(self.page_7)

        #第八页，外协件信息查询-------------------------------------------------------------------------------
        self.page_8 = QtWidgets.QWidget()
        self.page_8.setObjectName("page_8")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.page_8)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.gridLayout_5 = QtWidgets.QGridLayout()
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.exQMetal = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQMetal, 'metalForming')
        
        self.exQMetal.setMinimumSize(QtCore.QSize(200, 0))
        self.exQMetal.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exQMetal.setObjectName("exQMetal")
        self.gridLayout_5.addWidget(self.exQMetal, 5, 3, 1, 1)
        self.label_104 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_104.setFont(font)
        self.label_104.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_104.setObjectName("label_104")
        self.gridLayout_5.addWidget(self.label_104, 2, 0, 1, 1)
        self.label_36 = QtWidgets.QLabel(self.page_8)
        self.label_36.setMaximumSize(QtCore.QSize(300, 16777215))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_36.setFont(font)
        self.label_36.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_36.setObjectName("label_36")
        self.gridLayout_5.addWidget(self.label_36, 0, 0, 1, 1)
        self.exQInjection = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQInjection, 'injection')

        self.exQInjection.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exQInjection.setObjectName("exQInjection")
        self.gridLayout_5.addWidget(self.exQInjection, 4, 1, 1, 1)
        self.label_102 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_102.setFont(font)
        self.label_102.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_102.setObjectName("label_102")
        self.gridLayout_5.addWidget(self.label_102, 1, 0, 1, 1)
        self.label_106 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_106.setFont(font)
        self.label_106.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_106.setObjectName("label_106")
        self.gridLayout_5.addWidget(self.label_106, 3, 0, 1, 1)
        self.label_123 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_123.setFont(font)
        self.label_123.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_123.setObjectName("label_123")
        self.gridLayout_5.addWidget(self.label_123, 4, 0, 1, 1)
        self.exQProNumCombo = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQProNumCombo, 'projectNumber')

        self.exQProNumCombo.setObjectName("exQProNumCombo")
        self.gridLayout_5.addWidget(self.exQProNumCombo, 2, 1, 1, 1)
        self.exQProNameCombo = QtWidgets.QComboBox(self.page_8)

        #补全
        #self.exQProNameCombo.setCompleter(self.completer)
        #连接选项函数
        self.combo_option(self.exQProNameCombo, 'projectName')
        self.exQProNameCombo.setObjectName("exQProNameCombo")
        self.gridLayout_5.addWidget(self.exQProNameCombo, 1, 1, 1, 1)

        self.exQQEMcombo = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQQEMcombo, 'oem')

        #尝试补全
        tryList = ['张三','李四','王五']
        self.completer = QCompleter(tryList)
        self.completer.setFilterMode(Qt.MatchContains)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        self.exQQEMcombo.setCompleter(self.completer)

        self.exQQEMcombo.setMinimumSize(QtCore.QSize(200, 0))
        self.exQQEMcombo.setObjectName("exQQEMcombo")
        self.gridLayout_5.addWidget(self.exQQEMcombo, 0, 1, 1, 1)
        self.label_125 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_125.setFont(font)
        self.label_125.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_125.setObjectName("label_125")
        self.gridLayout_5.addWidget(self.label_125, 5, 0, 1, 1)
        self.exQDecorate = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQDecorate, 'decoration')

        self.exQDecorate.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exQDecorate.setObjectName("exQDecorate")
        self.gridLayout_5.addWidget(self.exQDecorate, 5, 1, 1, 1)
        self.exQTAcombo = QtWidgets.QComboBox(self.page_8)



        #连接选项函数
        self.combo_option(self.exQTAcombo, 'TAer')

        self.exQTAcombo.setObjectName("exQTAcombo")
        self.gridLayout_5.addWidget(self.exQTAcombo, 3, 1, 1, 1)
        self.label_127 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_127.setFont(font)
        self.label_127.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_127.setObjectName("label_127")
        self.gridLayout_5.addWidget(self.label_127, 6, 0, 1, 1)
        self.label_124 = QtWidgets.QLabel(self.page_8)
        self.label_124.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_124.setFont(font)
        self.label_124.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_124.setObjectName("label_124")
        self.gridLayout_5.addWidget(self.label_124, 4, 2, 1, 1)
        self.exQPart = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQPart, 'parts')

        self.exQPart.setMaximumSize(QtCore.QSize(200, 16777215))
        self.exQPart.setObjectName("exQPart")
        self.gridLayout_5.addWidget(self.exQPart, 6, 1, 1, 1)
        self.label_126 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_126.setFont(font)
        self.label_126.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_126.setObjectName("label_126")
        self.gridLayout_5.addWidget(self.label_126, 5, 2, 1, 1)
        self.exQClear = QtWidgets.QPushButton(self.page_8)

        #连接清除选项函数
        self.exQClear.clicked.connect(self.clear_all)

        self.exQClear.setMaximumSize(QtCore.QSize(80, 16777215))
        self.exQClear.setObjectName("exQClear")
        self.gridLayout_5.addWidget(self.exQClear, 7, 4, 1, 1)
        self.exQquery = QtWidgets.QPushButton(self.page_8)

        #连接查询函数
        self.exQquery.clicked.connect(self.external_execute_sql)

        self.exQquery.setMaximumSize(QtCore.QSize(80, 16777215))
        self.exQquery.setObjectName("exQquery")
        self.gridLayout_5.addWidget(self.exQquery, 7, 2, 1, 1)
        self.exQWield = QtWidgets.QComboBox(self.page_8)

        #连接选项函数
        self.combo_option(self.exQWield, 'welding')

        self.exQWield.setObjectName("exQWield")
        self.gridLayout_5.addWidget(self.exQWield, 4, 3, 1, 1)
        self.pushButton = QtWidgets.QPushButton(self.page_8)

        #连接导出函数
        self.pushButton.clicked.connect(self.external_query_result_to_excel)

        self.pushButton.setMaximumSize(QtCore.QSize(80, 16777215))
        self.pushButton.setObjectName("pushButton")
        self.gridLayout_5.addWidget(self.pushButton, 7, 3, 1, 1)
        self.verticalLayout_4.addLayout(self.gridLayout_5)
        self.exQuery = QtWidgets.QTableWidget(self.page_8)
        self.exQuery.setObjectName("exQuery")
        self.exQuery.setColumnCount(0)
        self.exQuery.setRowCount(0)
        self.verticalLayout_4.addWidget(self.exQuery)
        self.exQEdit = QtWidgets.QPushButton(self.page_8)
        self.exQEdit.setObjectName("exQEdit")
        self.verticalLayout_4.addWidget(self.exQEdit)
        self.exQClose = QtWidgets.QPushButton(self.page_8)

        #连接关闭函数
        self.exQClose.clicked.connect(self.q_close)
        self.exQClose.setObjectName("exQClose")
        self.verticalLayout_4.addWidget(self.exQClose)
        self.stackedWidget.addWidget(self.page_8)

        #补充控件用于自动补全
        self.qOemCompleter = QtWidgets.QLineEdit(self.page_8)

        #自动补全
        self.auto_complete_option( 'OEM')
        self.oemCompleter = QCompleter(self.oemList)
        #print(self.strList)
        self.oemCompleter.setFilterMode(Qt.MatchContains)
        self.oemCompleter.setCompletionMode(QCompleter.PopupCompletion)        
        self.qOemCompleter.setCompleter(self.oemCompleter)

        #连接控制选项函数
        #self.qOemCompleter.editingFinished.connect(self.select_by_auto_completer)
        self.qOemCompleter.textChanged.connect(lambda:self.select_by_auto_completer(self.qOemCompleter, self.exQQEMcombo))


        self.qOemCompleter.setMaximumSize(QtCore.QSize(200, 16777215))
        self.qOemCompleter.setObjectName("qOemCompleter")
        self.gridLayout_5.addWidget(self.qOemCompleter, 0, 3, 1, 1)

        self.qProNamCompleter = QtWidgets.QLineEdit(self.page_8)

        #自动补全函数
        self.auto_complete_option( 'projectName')
        #print(self.nameList)
        self.nameCompleter = QCompleter(self.nameList)
        self.nameCompleter.setFilterMode(Qt.MatchContains)
        self.nameCompleter.setCompletionMode(QCompleter.PopupCompletion)        
        self.qProNamCompleter.setCompleter(self.nameCompleter)

        #改变下拉框选项
        self.qProNamCompleter.textChanged.connect(lambda:self.select_by_auto_completer(self.qProNamCompleter, self.exQProNameCombo))

        self.qProNamCompleter.setMaximumSize(QtCore.QSize(400, 16777215))
        self.qProNamCompleter.setObjectName("qProNamCompleter")
        self.gridLayout_5.addWidget(self.qProNamCompleter, 1, 3, 1, 1)

        self.qProNumCompleter = QtWidgets.QLineEdit(self.page_8)

        #补全函数
        self.auto_complete_option( 'projectNumber')
        #print(self.nameList)
        self.numberCompleter = QCompleter(self.numberList)
        self.numberCompleter.setFilterMode(Qt.MatchContains)
        self.numberCompleter.setCompletionMode(QCompleter.PopupCompletion)        
        self.qProNumCompleter.setCompleter(self.numberCompleter)

        #改变下拉框选项
        self.qProNumCompleter.textChanged.connect(lambda:self.select_by_auto_completer(self.qProNumCompleter, self.exQProNumCombo))
        self.qProNumCompleter.setMaximumSize(QtCore.QSize(200, 16777215))
        self.qProNumCompleter.setObjectName("qProNumCompleter")
        self.gridLayout_5.addWidget(self.qProNumCompleter, 2, 3, 1, 1)
        self.label_101 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_101.setFont(font)
        self.label_101.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_101.setObjectName("label_101")
        self.gridLayout_5.addWidget(self.label_101, 0, 2, 1, 1)
        self.label_103 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_103.setFont(font)
        self.label_103.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_103.setObjectName("label_103")
        self.gridLayout_5.addWidget(self.label_103, 1, 2, 1, 1)
        self.label_105 = QtWidgets.QLabel(self.page_8)
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_105.setFont(font)
        self.label_105.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_105.setObjectName("label_105")
        self.gridLayout_5.addWidget(self.label_105, 2, 2, 1, 1)
        #self.verticalLayout_4.addLayout(self.gridLayout_5)

        #第九页，界面更改-----------------------------------------------------------------2019.8.25
        self.page_5 = QtWidgets.QWidget()
        self.page_5.setObjectName("page_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.page_5)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.gridLayout_8 = QtWidgets.QGridLayout()
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.optionImport = QtWidgets.QPushButton(self.page_5)

        #连接选项更新函数
        self.optionImport.clicked.connect(self.option_change)
        self.optionImport.setMaximumSize(QtCore.QSize(150, 16777215))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.optionImport.setFont(font)
        self.optionImport.setObjectName("optionImport")
        self.gridLayout_8.addWidget(self.optionImport, 1, 2, 1, 1)
        self.label_108 = QtWidgets.QLabel(self.page_5)
        self.label_108.setText("")
        self.label_108.setObjectName("label_108")
        self.gridLayout_8.addWidget(self.label_108, 0, 0, 1, 1)
        self.label_27 = QtWidgets.QLabel(self.page_5)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_27.setFont(font)
        self.label_27.setAlignment(QtCore.Qt.AlignCenter)
        self.label_27.setObjectName("label_27")
        self.gridLayout_8.addWidget(self.label_27, 0, 1, 1, 1)
        self.label_141 = QtWidgets.QLabel(self.page_5)
        self.label_141.setText("")
        self.label_141.setObjectName("label_141")
        self.gridLayout_8.addWidget(self.label_141, 0, 2, 1, 1)
        self.optionView = QtWidgets.QComboBox(self.page_5)

        self.optionView.addItem('')
        self.optionView.addItem('技术能力汇总')
        self.optionView.addItem('省份')
        self.optionView.addItem('注塑发泡类')
        self.optionView.addItem('表面装饰类')
        self.optionView.addItem('金属成型类')
        self.optionView.addItem('焊接类')
        self.optionView.addItem('重点零件类')
        self.optionView.addItem('主机厂')
        self.optionView.addItem('开发部门')
        self.optionView.addItem('TA工程师')

        #连接显示选项函数
        self.optionView.activated.connect(self.option_show)

        self.optionView.setMaximumSize(QtCore.QSize(200, 16777215))
        self.optionView.setObjectName("optionView")
        self.gridLayout_8.addWidget(self.optionView, 1, 1, 1, 1)
        self.optionTable = QtWidgets.QTableWidget(self.page_5)
        self.optionTable.setObjectName("optionTable")
        self.optionTable.setColumnCount(0)
        self.optionTable.setRowCount(0)
        self.gridLayout_8.addWidget(self.optionTable, 2, 0, 1, 3)
        self.label_140 = QtWidgets.QLabel(self.page_5)
        self.label_140.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("Adobe Devanagari")
        font.setPointSize(12)
        self.label_140.setFont(font)
        self.label_140.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_140.setObjectName("label_140")
        self.gridLayout_8.addWidget(self.label_140, 1, 0, 1, 1)
        self.verticalLayout_5.addLayout(self.gridLayout_8)
        self.label_28 = QtWidgets.QLabel(self.page_5)
        self.label_28.setText("")
        self.label_28.setObjectName("label_28")
        self.verticalLayout_5.addWidget(self.label_28)
        self.stackedWidget.addWidget(self.page_5)
        self.gridLayout.addWidget(self.stackedWidget, 1, 1, 1, 1)
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setText("")
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 0, 1, 1, 1)
        #更新界面时，应当将这句放在最新界面下---------------------------------更新界面关键语句-2019.8.11---------------------
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


        #供应商信息录入页，文本框table切换顺序
        Form.setTabOrder(self.supName, self.supAbbreviation)
        Form.setTabOrder(self.supAbbreviation, self.supProvince)
        Form.setTabOrder(self.supProvince, self.supCity)
        Form.setTabOrder(self.supCity, self.supEquip)
        Form.setTabOrder(self.supEquip, self.supModelRearch)
        Form.setTabOrder(self.supModelRearch, self.supTechRearch)
        Form.setTabOrder(self.supTechRearch, self.supBrief)
        Form.setTabOrder(self.supBrief, self.supTech1)
        Form.setTabOrder(self.supTech1, self.supTech2)
        Form.setTabOrder(self.supTech2, self.supTech3)
        Form.setTabOrder(self.supTech3, self.supTech4)
        Form.setTabOrder(self.supTech4, self.supTech5)
        Form.setTabOrder(self.supTech5, self.supOneAttach)
        Form.setTabOrder(self.supOneAttach, self.supOneView)
        #Form.setTabOrder(self.supMoreAttach, self.supOneView)
        Form.setTabOrder(self.supOneView, self.supMoreView)
        Form.setTabOrder(self.supMoreView, self.supClear)
        Form.setTabOrder(self.supClear, self.supBackUp)
        Form.setTabOrder(self.supBackUp, self.supInputDb)
        Form.setTabOrder(self.supInputDb, self.supQprovin)
        Form.setTabOrder(self.supQprovin, self.supQcity)
        Form.setTabOrder(self.supQcity, self.supQclear)
        Form.setTabOrder(self.supQclear, self.supQview)
        Form.setTabOrder(self.supQview, self.supQbrief)
        Form.setTabOrder(self.supQbrief, self.subInputView)
        Form.setTabOrder(self.subInputView, self.supQtech)
        Form.setTabOrder(self.supQtech, self.treeWidget)

        #初始化下拉框
        self.clear_all()
        #------------------------------------------初始化变量--------------------------------
        #初始化一个变量，判断是单个数据导入还是批量导入
        self.value = None
        #初始化附件路径，若没有选择路径则自动为空
        self.filePath = ''
        #用于供应商信息查询
        self.techText = ' '
        #------------------------------------------此函数用于控件文本信息设置-------------------------------------------------------------------------------------------

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "TA信息管理"))
        self.treeWidget.headerItem().setText(0, _translate("Form", "目录"))
        __sortingEnabled = self.treeWidget.isSortingEnabled()
       
       #
        #self.exTabox.setItemText(0, _translate("Form", "try"))
        self.treeWidget.setFont(self.font)
        self.treeWidget.setSortingEnabled(False)
        self.treeWidget.topLevelItem(0).setText(0, _translate("Form", "供应商信息管理"))
        self.treeWidget.topLevelItem(0).child(0).setText(0, _translate("Form", "供应商信息录入"))
        self.treeWidget.topLevelItem(0).child(1).setText(0, _translate("Form", "供应商信息查询"))
        self.treeWidget.topLevelItem(1).setText(0, _translate("Form", "外协件信息管理"))
        self.treeWidget.topLevelItem(1).child(0).setText(0, _translate("Form", "外协件信息录入"))
        self.treeWidget.topLevelItem(1).child(1).setText(0, _translate("Form", "外协件信息查询"))
        self.treeWidget.topLevelItem(2).setText(0, _translate("Form", "界面管理"))
        self.treeWidget.topLevelItem(2).child(0).setText(0, _translate("Form", "选项变更"))
        self.treeWidget.setSortingEnabled(__sortingEnabled)
        self.label.setText(_translate("Form", "TA信息管理系统"))
        self.label_3.setText(_translate("Form", "欢迎使用TA信息管理系统"))
        self.label_4.setText(_translate("Form", "目前所包含模块：供应商信息管理、外协件信息管理、界面管理"))
        self.label_7.setFont(self.font)
        self.label_7.setText(_translate("Form", "供应商名称："))
        #新增零件
        self.label_33.setFont(self.font)
        self.label_33.setText(_translate("Form", "零件："))
        #新增金属成型
        self.label_31.setFont(self.font)
        self.label_31.setText(_translate("Form", "金属成型类："))
        #新增焊接类
        self.label_32.setFont(self.font)
        self.label_32.setText(_translate("Form", "焊接类："))
        #新增表面装饰类
        self.label_30.setFont(self.font)
        self.label_30.setText(_translate("Form", "表面装饰类："))
        #新增注塑发泡类
        self.label_29.setFont(self.font)
        self.label_29.setText(_translate("Form", "注塑发泡类："))
        
        #第四页，外协件信息录入控件信息---------------------------
        self.label_46.setText(_translate("Form", "金属成型类："))
        self.label_79.setText(_translate("Form", "外协件TA信息录入"))
        self.label_42.setText(_translate("Form", "焊接类："))
        self.label_44.setText(_translate("Form", "注塑发泡类："))
        self.label_37.setText(_translate("Form", "OEM名称："))
        self.label_86.setText(_translate("Form", "项目号："))
        self.label_58.setText(_translate("Form", "TA评分："))
        self.label_52.setText(_translate("Form", "TA评分："))
        self.label_55.setText(_translate("Form", "TA评分："))
        self.label_40.setText(_translate("Form", "评估产品："))
        self.exBackUp.setText(_translate("Form", "备份"))
        self.label_62.setText(_translate("Form", "TA评分："))
        self.label_50.setText(_translate("Form", "TA评分："))
        self.label_38.setText(_translate("Form", "项目名称："))
        self.exProNumber.setText(_translate("Form", "项目号索引"))
        self.exInputDb.setText(_translate("Form", "保存到数据库"))
        self.label_45.setText(_translate("Form", "装饰类："))
        self.label_49.setText(_translate("Form", "所在地："))
        self.exSupDevQuaReview.setText(_translate("Form", "供应商开发质量评估"))
        self.exClear.setText(_translate("Form", "清空"))
        self.exAttach.setText(_translate("Form", "添加附件"))
        self.label_54.setText(_translate("Form", "所在地："))
        self.label_51.setText(_translate("Form", "所在地："))
        self.label_39.setText(_translate("Form", "开发部门："))
        self.label_57.setText(_translate("Form", "所在地："))
        self.exTechReview.setText(_translate("Form", "技术评估"))
        self.label_61.setText(_translate("Form", "所在地："))
        self.label_48.setText(_translate("Form", "模具评审日期："))
        self.label_47.setText(_translate("Form", "TA评审日期："))
        self.label_43.setText(_translate("Form", "主要工艺→零件类："))
        self.label_25.setText(_translate("Form", "布点供应商1："))
        self.label_53.setText(_translate("Form", "布点供应商2："))
        self.label_56.setText(_translate("Form", "布点供应商3："))
        self.label_60.setText(_translate("Form", "定点供应商："))
        self.label_59.setText(_translate("Form", "布点供应商4："))
        self.label_26.setText(_translate("Form", "导入《外协件前期技术管理工作计划表》"))
        self.label_85.setText(_translate("Form", "TA工程师："))

        self.exOutputExcel.setText(_translate("Form", "导出Excel"))

        self.label_9.setFont(self.font)
        self.label_9.setText(_translate("Form", "所在地："))
        self.label_12.setFont(self.font)
        self.label_12.setText(_translate("Form", "工艺能力："))
        self.label_14.setFont(self.font)
        self.label_14.setText(_translate("Form", "公司简介："))
        self.label_15.setFont(self.font)
        self.label_15.setText(_translate("Form", "公司详情："))
        self.supOneAttach.setText(_translate("Form", "添加附件"))
        #self.supMoreAttach.setText(_translate("Form", "批量添加"))
        self.label_23.setFont(self.font)
        self.label_23.setText(_translate("Form", "模具开发能力："))
        self.supModelRearch.setItemText(1, _translate("Form", "有"))
        self.supModelRearch.setItemText(2, _translate("Form", "无"))
        self.supModelRearch.setItemText(3, _translate("Form", "未知"))
        self.supInputDb.setText(_translate("Form", "保存到数据库"))
        self.supOneView.setText(_translate("Form", "添加"))
        self.supMoreView.setText(_translate("Form", "批量添加"))
        self.supClear.setText(_translate("Form", "清除"))
        self.supBackUp.setText(_translate("Form", "备份"))
        self.label_24.setFont(self.font)
        self.label_24.setText(_translate("Form", "技术开发能力："))
        self.supTechRearch.setItemText(1, _translate("Form", "有"))
        self.supTechRearch.setItemText(2, _translate("Form", "无"))
        self.supTechRearch.setItemText(3, _translate("Form", "未知"))
        self.label_10.setFont(self.font)
        self.label_10.setText(_translate("Form", "省"))
        self.label_11.setFont(self.font)
        self.label_11.setText(_translate("Form", "市"))
        self.label_13.setFont(self.font)
        self.label_13.setText(_translate("Form", "设备能力："))
        self.label_8.setFont(self.font)
        self.label_8.setText(_translate("Form", "简称："))
        self.label_16.setFont(self.font)
        self.label_16.setText(_translate("Form", "工艺或能力"))
        self.label_17.setFont(self.font)
        self.label_17.setText(_translate("Form", "供应商简称"))
        self.label_18.setFont(self.font)
        self.label_18.setText(_translate("Form", "所在省"))
        self.label_19.setFont(self.font)
        self.label_19.setText(_translate("Form", "所在市"))
        self.supQclear.setText(_translate("Form", "清除"))
        #self.label_25.setFont(self.font)
        #self.label_25.setText(_translate("Form", "功能2待开发"))
        self.label_27.setFont(self.font)
        self.label_27.setText(_translate("Form", "选项变更"))

        #补充日期格式提醒
        self.label_142.setText(_translate("Form", "格式：xxxx/xx/xx"))
        self.label_143.setText(_translate("Form", "格式：xxxx/xx/xx"))

        #第六页，技术评估报告--------------------------------
        self.label_34.setText(_translate("Form", "供应商名称："))
        self.label_80.setText(_translate("Form", "所在地："))
        self.label_81.setText(_translate("Form", "省"))
        self.label_82.setText(_translate("Form", "市"))
        self.label_83.setText(_translate("Form", "项目名："))
        self.label_87.setText(_translate("Form", "评估产品："))
        self.label_88.setText(_translate("Form", "主要工艺→零件类："))
        self.label_89.setText(_translate("Form", "注塑发泡类："))
        self.label_91.setText(_translate("Form", "金属成型类："))
        self.label_94.setText(_translate("Form", "评估得分："))
        self.exTechClear.setText(_translate("Form", "清空"))
        self.label_93.setText(_translate("Form", "评估项："))
        self.textBrowser.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">项目小组名单，成员各司其职，分工明确。（满分3分）</span></p></body></html>"))
        self.label_97.setText(_translate("Form", "项目团队能力："))
        self.textBrowser_2.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">以往项目的支持和配合程度。（满分10分）</span></p></body></html>"))
        self.label_99.setText(_translate("Form", "目标产品开发经验："))
        self.textBrowser_3.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">有类似关键零件或关键工艺的开发经验。（满分：35分）</span></p></body></html>"))
        self.label_100.setText(_translate("Form", "《技术协议》的理解："))
        self.textBrowser_4.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">输出问题清单。（满分：5分）</span></p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Form", "Tab 1"))
        self.textBrowser_5.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">数据的理解，输出“产品/模具可行性分析报告”。（满分：30分，低于10分一票否决）</span></p></body></html>"))
        self.textBrowser_6.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">图纸的理解。理解基准和关键尺寸要求，以及输出特殊特性检测和工装的方案。（满分：5分）</span></p></body></html>"))
        self.textBrowser_7.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">标准的理解。输出PV实验大纲，及相关风险问题。（满分：5分）</span></p></body></html>"))
        self.textBrowser_8.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">BOM的理解。确认BOM中的原材料、标准等信息。（满分：5分）</span></p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("Form", "Tab 2"))
        self.textBrowser_9.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">输出目标产品工艺流程图。定义了各工序的设备/设施、工装、模具、检具：输出对应特殊性清单。（满分：10分）</span></p></body></html>"))
        self.textBrowser_10.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">输出目标产品的：模具清单、检具清单、设备/工装/设施清单。（满分：6分）</span></p></body></html>"))
        self.textBrowser_11.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">输出PFMEA（满分：3分）</span></p></body></html>"))
        self.textBrowser_12.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">输出供应商的开发进度计划（满分：3分）</span></p></body></html>"))
        self.textBrowser_13.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">方案评估可信，每一项加2分。</span></p></body></html>"))
        self.label_109.setText(_translate("Form", "可能的WAVE方案："))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("Form", "Tab 3"))
        self.exTechImport.setText(_translate("Form", "从excel导入"))
        self.exTechInputDb.setText(_translate("Form", "保存到数据库"))
        self.label_110.setText(_translate("Form", "总分："))
        self.exTechBack.setText(_translate("Form", "返回"))
        self.label_95.setText(_translate("Form", "评估结果简述（T2前）"))
        self.label_90.setText(_translate("Form", "表面装饰类："))
        self.label_92.setText(_translate("Form", "焊接类："))
        self.label_63.setText(_translate("Form", "布点供应商编号："))
        self.exTechSupNum.setItemText(0, _translate("Form", " "))
        self.exTechSupNum.setItemText(1, _translate("Form", "1"))
        self.exTechSupNum.setItemText(2, _translate("Form", "2"))
        self.exTechSupNum.setItemText(3, _translate("Form", "3"))
        self.exTechSupNum.setItemText(4, _translate("Form", "4"))
        self.label_34.setText(_translate("Form", "供应商名称："))
        self.label_96.setText(_translate("Form", "评估结果简述（T2后）"))
        self.label_107.setText(_translate("Form", "评审时间："))
        self.exTechCal.setText(_translate("Form", "计算"))

        #第七页供应商质量评估报告----------------------------------------------------------
        self.label_115.setText(_translate("Form", "焊接类："))
        self.label_112.setText(_translate("Form", "注塑发泡类："))
        self.label_114.setText(_translate("Form", "表面装饰类："))
        self.label_120.setText(_translate("Form", "项目号："))
        self.label_118.setText(_translate("Form", "市"))
        self.label_116.setText(_translate("Form", "所在地："))
        self.label_35.setText(_translate("Form", "供应商名称："))
        self.label_117.setText(_translate("Form", "省"))
        self.label_119.setText(_translate("Form", "项目名："))
        self.label_84.setText(_translate("Form", "评估产品："))
        self.label_69.setText(_translate("Form", "技术能力："))
        self.textBrowser_14.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">关键技术能力：问题分析的能力；对性能要求的理解能力；模具方案的快速修改能力（满分：25分）</span></p></body></html>"))
        self.textBrowser_15.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">其他技术能力：手工样件验证能力；结构优化数据设计；跟踪产品装车验证；对二级供应商质量管控能力（满分：10分）</span></p></body></html>"))
        self.textBrowser_16.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">递交物质量：技术文件（如bom、实验报告等）；认可样品（ESO、PPAP）；验证用样品（满分：15分）</span></p></body></html>"))
        self.label_121.setText(_translate("Form", "合计："))
        self.exSupCal.setText(_translate("Form", "计算"))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_4), _translate("Form", "Tab 1"))
        self.label_70.setText(_translate("Form", "支持配合："))
        self.textBrowser_17.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">人员稳定性情况（满分：10分）</span></p></body></html>"))
        self.textBrowser_18.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">解决问题的响应速度（满分：10分）</span></p></body></html>"))
        self.textBrowser_19.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">样品递交的及时性（满分：10分）</span></p></body></html>"))
        self.textBrowser_20.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">其他及时性：文件递交的及时性；产品装车问题的应对速度；项目认可的及时性（满分：10分）</span></p></body></html>"))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_5), _translate("Form", "Tab 2"))
        self.textBrowser_21.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">供应商诚信度问题（如工艺环节内部简化、擅自更换原材料、递交文件作假等）。将取消下个项目的TA资格。如有，则需提供具体的案例和证据。</span></p></body></html>"))
        self.label_78.setText(_translate("Form", "一票否决问题："))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_6), _translate("Form", "Tab 3"))
        self.label_113.setText(_translate("Form", "金属成型类："))
        self.label_41.setText(_translate("Form", "类别："))
        self.label_111.setText(_translate("Form", "主要工艺→零件类："))
        self.exQuaImport.setText(_translate("Form", "导入excel"))
        self.exQuaInputDb.setText(_translate("Form", "保存到数据库"))
        self.exQuaClear.setText(_translate("Form", "清空"))
        self.exTechBack_2.setText(_translate("Form", "返回"))
        self.label_67.setText(_translate("Form", "具体描述："))
        self.label_68.setText(_translate("Form", "评估得分:"))

#第八页，查询页------------------------------------------------------------------------
        self.label_104.setText(_translate("Form", "项目号："))
        self.label_36.setText(_translate("Form", "OEM名称："))
        self.label_102.setText(_translate("Form", "项目名称："))
        self.label_106.setText(_translate("Form", "TA工程师："))
        self.label_123.setText(_translate("Form", "主要工艺-注塑发泡类："))
        self.label_125.setText(_translate("Form", "表面装饰类："))
        self.label_127.setText(_translate("Form", "重点零件类："))
        self.label_124.setText(_translate("Form", "金属成型类："))
        self.label_126.setText(_translate("Form", "焊接类："))
        self.exQClear.setText(_translate("Form", "清除"))
        self.exQquery.setText(_translate("Form", "查询"))
        self.pushButton.setText(_translate("Form", "导出"))
        self.exQEdit.setText(_translate("Form", "编辑"))
        self.exQClose.setText(_translate("Form", "关闭"))

        self.label_101.setText(_translate("Form", "或："))
        self.label_103.setText(_translate("Form", "或："))
        self.label_105.setText(_translate("Form", "或："))

        #第九页，选项更改
        self.optionImport.setText(_translate("Form", "导入更改"))
        self.label_27.setText(_translate("Form", ""))
        self.label_140.setText(_translate("Form", "选项查看："))


        #2019.8.25---------------------------------选项更改界面

        #更改选项按钮
    def option_change(self):

        #连接表格
        try:
            book = xlrd.open_workbook('resourse\\下拉选项变更模板.xlsx')
        except:
            QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        try:
            sheet = book.sheet_by_name("Sheet1")
        except:
            QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


        #提取excel数据

        #存放excel中一行数据
        excelRow = []
        #包含所有excel行数据
        excelData = []
        for i in range(1, sheet.nrows):
            #存完一行后及时清空
            excelRow = []
            for j in range(0, sheet.ncols):
                excelRow.append(sheet.cell(i,j).value)
            excelData.append(list(excelRow))

        #print(str(len(excelData)))
        #print(str(sheet.nrows))
        #print(str(sheet.ncols))
        #print(excelData[0])
        #print(excelData[2])

        #导入数据库
        try:
            str01 = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str01)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


        sql = "update ConstantOption set technologicalCapability = ?, province = ?, injection = ?, \
        decoration = ?, metalForming = ?, welding = ?, parts = ?, oem = ?, depName = ? , TAer = ?  where [ID] = ?"

        para = []
        for i in range(len(excelData)):
            #print(excelData[i][0])
            para = [excelData[i][0], excelData[i][1], excelData[i][2], excelData[i][3], excelData[i][4], excelData[i][5],excelData[i][6], excelData[i][7], excelData[i][8], excelData[i][9], str(i + 1)]
            cursor.execute(sql, para)
            cursor.commit()
            para = []

        #更新选型
        #self.combo_option(self.supTech1, 'parts')
        QMessageBox.about(self, '提示', '修改完成！')

        cursor.close()
        db.close()





        #显示选项
    def option_show(self):
        #连接数据库
        try:
            str01 = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str01)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


        sql = "select * from ConstantOption "
        #print(sql)
        cursor.execute(sql)
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
        rows = cursor.fetchall()
        
        nrows = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)

        #print(nrows[0][0])
        #print(nrows[0][1])

        #text = self.optionView.currentText()
        #返回对应选项列编号
        colNumber = self.optionView.currentIndex()
        #QMessageBox.about(self, '提示', str(colNumber))
        #获取对应选项个数
        colLength = 0 
        #QMessageBox.about(self, '提示', str(colLength))
        for i in range(len(nrows)):
            if colNumber == 1:
                if nrows[i][colNumber] != None:
                    colLength = i + 1
                    #QMessageBox.about(self, '提示', str(colLength))
            else:
                if nrows[i][colNumber] != '' and nrows[i][colNumber] != None:
                    colLength = i + 1
                    #QMessageBox.about(self, '提示', str(colLength))
                #QMessageBox.about(self, '提示', nrows[i][colNumber])
        #if nrows[6][2] == '':
            #QMessageBox.about(self, '提示', '为空')
        #QMessageBox.about(self, '提示', str(colLength))
        #设置行列
        self.optionTable.setColumnCount(1)
        self.optionTable.setRowCount(colLength)
        self.optionTable.setHorizontalHeaderLabels(['选项']) 
        #不可编辑 
        self.optionTable.setEditTriggers(QTableView.NoEditTriggers)
        #自适应窗口大小
        self.optionTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
#添加查询结果
        for i in range(colLength):
            
            
            item = QTableWidgetItem(" %s" % nrows[i][colNumber])
            #设置字体大小
            item.setFont(self.font)
            item.setTextAlignment(Qt.AlignCenter| Qt.AlignCenter)
            
            self.optionTable.setItem(i-1, 1, item)
        cursor.close() 
        db.close()  

        #2019.8.18--------------------------------外协件信息查询界面

        #导出查询结果到excel
    def external_query_result_to_excel(self):
        #将模板excel复制到用户指定位置
        directory1 = QFileDialog.getExistingDirectory(self,
                                    "选取文件夹",
                                    "C:/")    
        nowtime=time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))
        backPath = directory1  +  '\\外协件信息查询结果-' + nowtime  + '.xlsx'
        #print(backPath)
        original = os.getcwd()  + r'\resourse\\' + '外协件信息查询结果模板.xlsx'
        #复制并重命名
        shutil.copy(original, backPath)

        #填充表格
        #openpyxl连接表格
        try:
            book = openpyxl.load_workbook(backPath)
        except:
            QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #xlwt连接sheet,行列号起始数皆为0
        try:
            #ws = book.get_sheet_by_name("sheet1")
            ws = book.worksheets[0]
        except:
            QMessageBox.about(self, '错误', 'sheet定位失败！')

        #连接数据库
        try:
            stra = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(stra)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


        sql = "select * from External " + self.external_sql()
        print(sql)
        
        cursor.execute(sql)
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
        rows = cursor.fetchall()
        #print(rows)
        #nrows[0] = '产品名称'，nrows[0] = 'TA评审时间，及外协件评审时间'，nrows[0] = '模具评审时间'
        nrows = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)
        #excel赋值不接受None，如果不存在赋予' '
        i = 3
        for row in nrows:
            for j in range(194):
                if row[j] is None :
                    row[j] = str(row[j])
                else:
                    row[j] = row[j]
                #print(row[j])
                #这地方是个大坑，报错索引至少为1，参照 https://blog.csdn.net/Linli522362242/article/details/97469795
                ws.cell(row = i + 1, column = j + 1).value =nrows[i-3][j]
            i = i + 1
            #print(i)
        #print(nrows)
           
        
                
            
            
   
        book.save(filename = backPath)
        QMessageBox.about(self, '提示', '导出完成！')
         #断开连接
        cursor.close()
        db.close()



        #根据不同选项获得sql语句
    def external_sql(self):
        #获取关键字文本
        oem = self.exQQEMcombo.currentText()
        name = self.exQProNameCombo.currentText()
        number = self.exQProNumCombo.currentText()
        ta = self.exQTAcombo.currentText()
        inje = self.exQInjection.currentText()
        dec = self.exQDecorate.currentText()
        part = self.exQPart.currentText()
        wield = self.exQWield.currentText()
        metal = self.exQMetal.currentText()

        #一种情况
        #oem
        if oem != '' and name == '' and number == '' and ta == '' and inje == '' and dec == '' and part == '' and wield == '' and metal == ''  '':   
            #self.exQProNameCombo.setCurrentIndex(-1)
            #self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'oem！')
            sql02 = "where OEM like '{0}'".format(oem)
            #self.external_execute_sql(sql02)

        #name
        elif name !='' and oem =='':
            #self.exQQEMcombo.setCurrentIndex(-1)
            #self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'name！')
            sql02 = "where projectName like '{0}'".format(name)
            #self.external_execute_sql(sql02)

        #number
        elif number !='' and oem=='':
            #self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            #self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'number！')
            sql02 = "where projectNumber like '{0}'".format(number)
            #self.external_execute_sql(sql02)

        #TAer
        elif ta != '' and oem=='' :
            self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            #self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'ta！')
            sql02 = "where exTa like '{0}'".format(ta)
            #self.external_execute_sql(sql02)

        #injection
        elif inje != '' and oem=='':
            self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            #self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'inje！')
            sql02 = "where exInjection like '{0}'".format(inje)
            #self.external_execute_sql(sql02)

        #decration
        elif dec != '' and oem=='':
            self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            #self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'dec！')
            sql02 = "where exDecorate like '{0}'".format(dec)
            #self.external_execute_sql(sql02)

        #part
        elif part != '' and oem=='':
            self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            #self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'part！')
            sql02 = "where exPart like '{0}'".format(part)
            #self.external_execute_sql(sql02)

        #wield
        elif wield != '' and oem=='':
            self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            #self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'wield！')
            sql02 = "where exWield like '{0}'".format(wield)
            #self.external_execute_sql(sql02)

        #metal
        elif metal != '' and oem=='':
            self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            #self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'metal！')
            sql02 = "where exMetal like '{0}'".format(metal)
            #self.external_execute_sql(sql02)


        #两种
        #oem and name
        if oem != '' and name != '' and number == '' :
            #self.exQQEMcombo.setCurrentIndex(-1)
            #self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'oem and name！')
            sql02 = "where OEM like '{0}' and ( projectName like '{1}')".format(oem, name)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            #self.execute_sql(sql02)

        #oem and number
        elif oem != '' and name == '' and number != '' :
            #self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            #self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'oem and number！')
            sql02 = "where OEM like '{0}' and ( projectNumber like '{1}')".format(oem, number)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            #self.execute_sql(sql02)

            #oem name number都有默认查询oem and name
        elif  oem != '' and name != '' and number != '' :
            #self.exQQEMcombo.setCurrentIndex(-1)
            #self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            QMessageBox.about(self, '提示', 'oem and name！')
            sql02 = "where OEM like '{0}' and ( projectName like '{1}')".format(oem, name)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            #self.execute_sql(sql02)

        #其余所有情况只查询oem
        else:
            #self.exQQEMcombo.setCurrentIndex(-1)
            self.exQProNameCombo.setCurrentIndex(-1)
            self.exQProNumCombo.setCurrentIndex(-1)
            self.exQTAcombo.setCurrentIndex(-1)
            self.exQInjection.setCurrentIndex(-1)
            self.exQDecorate.setCurrentIndex(-1)
            self.exQPart.setCurrentIndex(-1)
            self.exQWield.setCurrentIndex(-1)
            self.exQMetal.setCurrentIndex(-1)
            #QMessageBox.about(self, '提示', 'oem ！')
            sql02 = "where OEM like '{0}'".format(oem)
            #self.external_execute_sql(sql02)

        return sql02

        #执行sql语句
    def external_execute_sql(self):
        self.external_sql()
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        #含布点供应商
        #sql = " select OEM, projectName, projectNumber, department, productName, exTa, exPart, exInjection, exMetal, exDecorate, exWield, \
        #externalTime, mouldTime, exSeleSup, exSeleSupPlace , exSeleSupTaMark, distributionSup1,exSup1Place, sup1Mak, distributionSup2, \
        #exSup2Place, sup2Mak, distributionSup3, exSup3Place, sup3Mak, distributionSup4,  exSup4Place, sup4Mak from External " + sql02

        sql = " select OEM, projectName, projectNumber, department, productName, exTa, exPart, exInjection, exMetal, exDecorate, exWield, \
        externalTime, mouldTime, exSeleSup, exSeleSupPlace , exSeleSupTaMark from External " + self.external_sql()

        #print(sql)
        #QMessageBox.information(self, '提示', sql, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        cursor.execute(sql)
        
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
        rows = cursor.fetchall()
        nrows = []
        path = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)

         #设置行列
        self.exQuery.setColumnCount(16)
        self.exQuery.setRowCount(len(nrows))
        self.exQuery.setHorizontalHeaderLabels(['主机厂', '项目名称','项目号', '开发部门', '产品名称', 'TA工程师','零件类', 
            '注塑发泡类', '金属成型类', '装饰类', '焊接类', 'TA评审日期', '模具评审日期', '定点供应商','所在地', 'TA评分']) 
        #不可编辑 
        #self.exQuery.setEditTriggers(QTableView.NoEditTriggers)
        self.exQuery.setEditTriggers(QAbstractItemView.DoubleClicked)
        #自适应窗口大小，与根据内容调整列宽冲突
        #self.exQuery.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        #选取策略整行
        #self.exQuery.setSelectionBehavior(QAbstractItemView.SelectRows)
        #根据内容调整列宽
        QTableWidget.resizeColumnsToContents(self.exQuery)

        #填充内容
        for row in range(len(nrows)):
            for column in range(16):
                #rowslice获取nrows中一行的数据，与获取path的rowSlice不同
                rowslice = nrows[row]
                item = QTableWidgetItem(" %s" % rowslice[column])
                self.exQuery.setItem(row, column, item)
        cursor.close() 
        db.close()  


        #关闭按钮
    def q_close(self):
        self.stackedWidget.setCurrentIndex(-1)
        #根据lineEdit补全的内容返回对应下拉框选项
    def select_by_auto_completer(self, lineEdit, box):
        if box.findText(lineEdit.text()):
           box.setCurrentText(lineEdit.text())


        #自动补全模块,返回自动补全的列表，语句需根据控件另外创建

    def auto_complete_option(self, keyWord):
        #--------------------采用pyodbc连接数据库，为下拉框提供选项------------------------------------------
        try:
            stra = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(stra)
            cursor = db.cursor()
        
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #科技能力和省采用ConstanOption表中选项，简称和市选用Supplier表
        
        if keyWord == 'projectName' or keyWord == 'projectNumber' or keyWord == 'OEM':
            sql = "select " + keyWord + " from External"
        
        #QMessageBox.information(self, '提示', sql,QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        cursor.execute(sql)
        rows = cursor.fetchall()
        nrows = []
        for row in rows:
            #无需int化，item只接受str类型
            #lRow = [int_it(each) for each in row]
            nrows.append(row)
        #去除重复项
        dNrows = []
        for i in nrows:
            if i not in dNrows :
                dNrows.append(i)

        #此列表用于储存补全数据来源
        if keyWord == 'projectName':
            self.nameList = []
            for i in range(len(dNrows)):
                self.nameList.append(str(dNrows[i][0]))
        elif keyWord == 'projectNumber':
            self.numberList = []
            for i in range(len(dNrows)):
                self.numberList.append(str(dNrows[i][0]))
        elif keyWord == 'OEM':
            self.oemList = []
            for i in range(len(dNrows)):
                self.oemList.append(str(dNrows[i][0]))


        
        


        #2019.8.16-供应商质量评估界面-------------------------------------

      

        #从excel导入-供应商质量评估，2019.8.25


    def exQua_import_excel(self):
        fileName = ''
        fileName, fileType = QFileDialog.getOpenFileName(self,
                                    "选取文件",
                                    os.getcwd() + "\\resourse\\appendix",
                                    "All Files (*);;Text Files (*.txt)")
        #批量导入
        #连接表格
        if fileName != '':
            try:
                book = xlrd.open_workbook(fileName)
            except:
                QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            try:
                sheet = book.sheet_by_name("反馈表")
            except:
                QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            mark = []
            #大坑,xlrd索引有时会报错超出范围，-1后调整循环参数即可
            for i in range(7,16):
                a = sheet.cell(i - 1 , 2).value
                b = str(a)
                mark.append(b)
            print(mark)
            self.exQuaM1.setText(mark[0])
            self.exQuaM2.setText(mark[1])
            self.exQuaM3.setText(mark[2])
            self.exQuaM4.setText(mark[3])
            self.exQuaM5.setText(mark[4])
            self.exQuaM6.setText(mark[5])
            self.exQuaM7.setText(mark[6])
            self.exQuaM8.setText(mark[7])
            self.exQuaTotal.setText(mark[8])

            QMessageBox.about(self, '提示', '导入完成！')
        else:
            QMessageBox.about(self, '提示', '未选择附件！')
        #导入数据库------------------------------------------------2019.8.16
        #连接数据库
    def qua_to_db(self):
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        #获取界面输入值
        m1 = self.exQuaM1.text()
        m2 = self.exQuaM2.text()
        m3 = self.exQuaM3.text()
        m4 = self.exQuaM4.text()
        m5 = self.exQuaM5.text()
        m6 = self.exQuaM6.text()
        m7 = self.exQuaM7.text()
        m8 = self.exQuaM8.text()
        m9 = self.exQuaTotal.text()

        sql01 = "update External set selectKeyTechMark = ?, selectOtherTechMark = ?, selectSamQuaMark = ?, \
        selectPerStabMark = ?, selectResponSdMark = ?, selecSampleTimeMark = ?,\
        selectOtherTimeMark = ?, selectSinMark = ?, selectQuaTotalMark = ? "
        sql02 = " where projectName = ? and [productName] = ? "
        
        sql = sql01 + sql02

        para = (m1, m2, m3, m4, m5, m6, m7, m8, m9, self.exProName.text(), self.exInputOption.currentText())

        cursor.execute(sql, para)
        cursor.commit()
        QMessageBox.about(self, '提示', '导入完成！')

        cursor.close()
        db.close()
        #计算按钮
    def cal_qua_mark(self):
        intMark = []
        intMark.append(self.exQuaM1.text())
        intMark.append(self.exQuaM2.text())
        intMark.append(self.exQuaM3.text())
        intMark.append(self.exQuaM4.text())
        intMark.append(self.exQuaM5.text())
        intMark.append(self.exQuaM6.text())
        intMark.append(self.exQuaM7.text())
        intMark.append(self.exQuaM8.text())
        
        mark = 0
        for i in range (len(intMark)):
            if intMark[i] == '':
                intMark[i] = 0
            mark = mark + int(intMark[i])
        

        self.exQuaTotal.setText(str(mark))

        #返回按钮
    def qua_back_tech(self):
        self.stackedWidget.setCurrentIndex(3)

        #--------------------2019.8.11-技术评估界面相关------------------------------

        #计算总分数
    def cal_tech_mark(self):
        intMark = []
        intMark.append(self.exTechM1.text())
        intMark.append(self.exTechM2.text())
        intMark.append(self.exTechM3.text())
        intMark.append(self.exTechM4.text())
        intMark.append(self.exTechM5.text())
        intMark.append(self.exTechM6.text())
        intMark.append(self.exTechM7.text())
        intMark.append(self.exTechM8.text())
        intMark.append(self.exTechM9.text())
        intMark.append(self.exTechM10.text())
        intMark.append(self.exTechM11.text())
        intMark.append(self.exTechM12.text())
        intMark.append(self.exTechM13.text())
        
        
        x = []
        mark = 0
        for i in range (len(intMark)):
            if intMark[i] == '':
                intMark[i] = 0
            mark = mark + int(intMark[i])
        

        self.exTechTotalMark.setText(str(mark))



        #从excel导入到界面
    def teche_review_from_excel(self):
        fileName = ''
        fileName, fileType = QFileDialog.getOpenFileName(self,
                                    "选取文件",
                                    os.getcwd() + "\\resourse\\appendix",
                                    "All Files (*);;Text Files (*.txt)")
        #批量导入
        #连接表格
        if fileName != '':
            try:
                book = xlrd.open_workbook(fileName)
            except:
                QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            try:
                sheet = book.sheet_by_name("评估报告")
            except:
                QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            #QMessageBox.information(self, '错误', fileName,QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            #赋值
            #分数
            mark = []
            for i in range(7,21):
                a = sheet.cell(i,5).value
                b = str(a)
                mark.append(b)

            self.exTechM1.setText(mark[0])
            self.exTechM2.setText(mark[1])
            self.exTechM3.setText(mark[2])
            self.exTechM4.setText(mark[0])
            self.exTechM5.setText(mark[3])
            self.exTechM6.setText(mark[4])
            self.exTechM7.setText(mark[5])
            self.exTechM8.setText(mark[6])
            self.exTechM9.setText(mark[7])
            self.exTechM10.setText(mark[8])
            self.exTechM11.setText(mark[9])
            self.exTechM12.setText(mark[10])
            self.exTechM13.setText(mark[11])
            self.exTechTotalMark.setText(mark[12])

            #Qtextedit
            #T2前
            self.exTechDesF1.setPlainText(sheet.cell(7,6).value)
            self.exTechDesF2.setPlainText(sheet.cell(8,6).value)
            self.exTechDesF3.setPlainText(sheet.cell(9,6).value)
            self.exTechDesF4.setPlainText(sheet.cell(10,6).value)
            self.exTechDesF5.setPlainText(sheet.cell(11,6).value)
            self.exTechDesF6.setPlainText(sheet.cell(12,6).value)
            self.exTechDesF7.setPlainText(sheet.cell(13,6).value)
            self.exTechDesF8.setPlainText(sheet.cell(14,6).value)
            self.exTechDesF9.setPlainText(sheet.cell(15,6).value)
            self.exTechDesF10.setPlainText(sheet.cell(16,6).value)
            self.exTechDesF11.setPlainText(sheet.cell(17,6).value)
            self.exTechDesF12.setPlainText(sheet.cell(18,6).value)
            self.exTechDesF13.setPlainText(sheet.cell(19,6).value)
            #T2后
            self.exTechDesA1.setPlainText(sheet.cell(7,7).value)
            self.exTechDesA2.setPlainText(sheet.cell(8,7).value)
            self.exTechDesA3.setPlainText(sheet.cell(9,7).value)
            self.exTechDesA4.setPlainText(sheet.cell(10,7).value)
            self.exTechDesA5.setPlainText(sheet.cell(11,7).value)
            self.exTechDesA6.setPlainText(sheet.cell(12,7).value)
            self.exTechDesA7.setPlainText(sheet.cell(13,7).value)
            self.exTechDesA8.setPlainText(sheet.cell(14,7).value)
            self.exTechDesA9.setPlainText(sheet.cell(15,7).value)
            self.exTechDesA10.setPlainText(sheet.cell(16,7).value)
            self.exTechDesA11.setPlainText(sheet.cell(17,7).value)
            self.exTechDesA12.setPlainText(sheet.cell(18,7).value)
            self.exTechDesA13.setPlainText(sheet.cell(19,7).value)

            QMessageBox.about(self, '提示', '导入完成！')
        else:
            QMessageBox.about(self, '提示', '未选择附件！')

        #导入数据库，根据下拉框判断是布点供应商几，采用insert into .. where 项目名+产品名来锁定外协件界面已导入的数据
    def tech_review_to_db(self):
        #连接数据库
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        #获取界面输入值
        m1 = self.exTechM1.text()
        m2 = self.exTechM2.text()
        m3 = self.exTechM3.text()
        m4 = self.exTechM4.text()
        m5 = self.exTechM5.text()
        m6 = self.exTechM6.text()
        m7 = self.exTechM7.text()
        m8 = self.exTechM8.text()
        m9 = self.exTechM9.text()
        m10 = self.exTechM10.text()
        m11 = self.exTechM11.text()
        m12 = self.exTechM12.text()
        m13 = self.exTechM13.text()
        #总分需要连接到外协件界面exSup1TaMark，并导入数据库
        mTotal = self.exTechTotalMark.text()

        #评估结果简述，T2前
        f1 = self.exTechDesF1.toPlainText()
        f2 = self.exTechDesF2.toPlainText()
        f3 = self.exTechDesF3.toPlainText()
        f4 = self.exTechDesF4.toPlainText()
        f5 = self.exTechDesF5.toPlainText()
        f6 = self.exTechDesF6.toPlainText()
        f7 = self.exTechDesF7.toPlainText()
        f8 = self.exTechDesF8.toPlainText()
        f9 = self.exTechDesF9.toPlainText()
        f10 = self.exTechDesF10.toPlainText()
        f11 = self.exTechDesF11.toPlainText()
        f12 = self.exTechDesF12.toPlainText()
        f13 = self.exTechDesF13.toPlainText()
        #评估结果简述，T2后
        a1 = self.exTechDesA1.toPlainText()
        a2 = self.exTechDesA2.toPlainText()
        a3 = self.exTechDesA3.toPlainText()
        a4 = self.exTechDesA4.toPlainText()
        a5 = self.exTechDesA5.toPlainText()
        a6 = self.exTechDesA6.toPlainText()
        a7 = self.exTechDesA7.toPlainText()
        a8 = self.exTechDesA8.toPlainText()
        a9 = self.exTechDesA9.toPlainText()
        a10 = self.exTechDesA10.toPlainText()
        a11 = self.exTechDesA11.toPlainText()
        a12 = self.exTechDesA12.toPlainText()
        a13 = self.exTechDesA13.toPlainText()

        #根据下拉框选项采用不同sql语句
        if self.exTechSupNum.currentText() == '1':
            self.exSup1TaMark.setText(mTotal)
            sql01 = "update External set sup1GroupMark = ?, sup1GroupSketch01 = ?, sup1GroupSketch02 = ?, sup1PastCoopMark = ?, sup1PastCoopSketch01 = ?, sup1PastCoopSketch02 = ?,\
sup1DevExpMark = ?, sup1DevExpSketch01 = ?, sup1DevExpSketch02 = ?, sup1ListProMark = ?, sup1ListProSketch01 = ?, sup1ListProSketch02 = ?, \
sup1DataUnderMark = ?, sup1DataUnderSketch01 = ?, sup1DataUnderSketch02 = ?, sup1DrawUnderMark = ?, sup1DrawUnderSketch01 = ?, sup1DrawUnderSketch02 = ?,\
sup1StandUnderMark = ?, sup1StandUnderSketch01 = ?, sup1StandUnderSketch02 = ?, sup1BomUnderMark = ?, sup1BomUnderSketch01 = ?, sup1BomUnderSketch02 = ?, \
sup1TechProMark = ?, sup1TechProSketch01 = ?, sup1TechProSketch02 = ?, sup1EquipMark = ?, sup1EquipSketch01 = ?, sup1EquipSketch02 = ?, \
sup1PfmeaMark = ?, sup1PfmeaSketch01 = ?, sup1PfmeaSketch02 = ?, sup1DevDesignMark = ?, sup1DevDesignSketch01 = ?, sup1DevDesignSketch02 = ?, sup1VaveMark = ?, sup1VaveSketch01 = ?, \
sup1VaveSketch02 = ?"
        elif self.exTechSupNum.currentText() == '2':
            self.exSup2TaMark.setText(mTotal)
            sql01 = sql01 = "update External set sup2GroupMark = ?, sup2GroupSketch01 = ?, sup2GroupSketch02 = ?, sup2PastCoopMark = ?, sup2PastCoopSketch01 = ?, sup2PastCoopSketch02 = ?,\
sup2DevExpMark = ?, sup2DevExpSketch01 = ?, sup2DevExpSketch02 = ?, sup2ListProMark = ?, sup2ListProSketch01 = ?, sup2ListProSketch02 = ?, \
sup2DataUnderMark = ?, sup2DataUnderSketch01 = ?, sup2DataUnderSketch02 = ?, sup2DrawUnderMark = ?, sup2DrawUnderSketch01 = ?, sup2DrawUnderSketch02 = ?,\
sup2StandUnderMark = ?, sup2StandUnderSketch01 = ?, sup2StandUnderSketch02 = ?, sup2BomUnderMark = ?, sup2BomUnderSketch01 = ?, sup2BomUnderSketch02 = ?, \
sup2TechProMark = ?, sup2TechProSketch01 = ?, sup2TechProSketch02 = ?, sup2EquipMark = ?, sup2EquipSketch01 = ?, sup2EquipSketch02 = ?, \
sup2PfmeaMark = ?, sup2PfmeaSketch01 = ?, sup2PfmeaSketch02 = ?, sup2DevDesignMark = ?, sup2DevDesignSketch01 = ?, sup2DevDesignSketch02 = ?, sup2VaveMark = ?, sup2VaveSketch01 = ?, \
sup2VaveSketch02 = ?"
        elif self.exTechSupNum.currentText() == '3':
            self.exSup3TaMark.setText(mTotal)
            sql01 = "update External set sup3GroupMark = ?, sup3GroupSketch01 = ?, sup3GroupSketch02 = ?, sup3PastCoopMark = ?, sup3PastCoopSketch01 = ?, sup3PastCoopSketch02 = ?,\
sup3DevExpMark = ?, sup3DevExpSketch01 = ?, sup3DevExpSketch02 = ?, sup3ListProMark = ?, sup3ListProSketch01 = ?, sup3ListProSketch02 = ?, \
sup3DataUnderMark = ?, sup3DataUnderSketch01 = ?, sup3DataUnderSketch02 = ?, sup3DrawUnderMark = ?, sup3DrawUnderSketch01 = ?, sup3DrawUnderSketch02 = ?,\
sup3StandUnderMark = ?, sup3StandUnderSketch01 = ?, sup3StandUnderSketch02 = ?, sup3BomUnderMark = ?, sup3BomUnderSketch01 = ?, sup3BomUnderSketch02 = ?, \
sup3TechProMark = ?, sup3TechProSketch01 = ?, sup3TechProSketch02 = ?, sup3EquipMark = ?, sup3EquipSketch01 = ?, sup3EquipSketch02 = ?, \
sup3PfmeaMark = ?, sup3PfmeaSketch01 = ?, sup3PfmeaSketch02 = ?, sup3DevDesignMark = ?, sup3DevDesignSketch01 = ?, sup3DevDesignSketch02 = ?, sup3VaveMark = ?, sup3VaveSketch01 = ?, \
sup3VaveSketch02 = ?"
        elif self.exTechSupNum.currentText() == '4':
            self.exSup4TaMark.setText(mTotal)
            sql01 = "update External set sup4GroupMark = ?, sup4GroupSketch01 = ?, sup4GroupSketch02 = ?, sup4PastCoopMark = ?, sup4PastCoopSketch01 = ?, sup4PastCoopSketch02 = ?,\
sup4DevExpMark = ?, sup4DevExpSketch01 = ?, sup4DevExpSketch02 = ?, sup4ListProMark = ?, sup4ListProSketch01 = ?, sup4ListProSketch02 = ?, \
sup4DataUnderMark = ?, sup4DataUnderSketch01 = ?, sup4DataUnderSketch02 = ?, sup4DrawUnderMark = ?, sup4DrawUnderSketch01 = ?, sup4DrawUnderSketch02 = ?,\
sup4StandUnderMark = ?, sup4StandUnderSketch01 = ?, sup4StandUnderSketch02 = ?, sup4BomUnderMark = ?, sup4BomUnderSketch01 = ?, sup4BomUnderSketch02 = ?, \
sup4TechProMark = ?, sup4TechProSketch01 = ?, sup4TechProSketch02 = ?, sup4EquipMark = ?, sup4EquipSketch01 = ?, sup4EquipSketch02 = ?, \
sup4PfmeaMark = ?, sup4PfmeaSketch01 = ?, sup4PfmeaSketch02 = ?, sup4DevDesignMark = ?, sup4DevDesignSketch01 = ?, sup4DevDesignSketch02 = ?, sup4VaveMark = ?, sup4VaveSketch01 = ?, \
sup4VaveSketch02 = ?"

        sql02 = " where projectName = ? and [productName] = ? "
        sql = sql01 + sql02
        para = (m1, f1, a1, m2, f2, a2, m3, f3, a3, m4, f4, a4, m5, f5, a5, m6, f6, a6,
             m7, f7, a7, m8, f8, a8, m9, f9, a9, m10, f10, a10, m11, f11, a11, m12, f12, a12,
              m13, f13, a13, self.exProName.text(), self.exInputOption.currentText())

        cursor.execute(sql, para)
        cursor.commit()
        QMessageBox.about(self, '提示', '导入完成！')

        cursor.close()
        db.close()

        


    def contentReference_from_external_to_techReciew(self):
            #供应商名称根据选项指定不同布点供应商
        if self.exTechSupNum.currentText() == '1':
            if self.exSup1.currentText() != ' ' :
                self.exTechSupName.setText(self.exSup1.currentText())
        elif self.exTechSupNum.currentText() == '2':
            if self.exSup2.currentText() != ' ' :
                self.exTechSupName.setText(self.exSup2.currentText())
        elif self.exTechSupNum.currentText() == '3':
            if self.exSup3.currentText() != ' ' :
                self.exTechSupName.setText(self.exSup3.currentText())
        elif self.exTechSupNum.currentText() == '4':
            if self.exSup4.currentText() != ' ' :
                self.exTechSupName.setText(self.exSup4.currentText())

            #项目名称，评估产品，评审时间，技术能力引用
        #项目名称
        if self.exProName.text() != ' ':
            self.exTechProName.setText(self.exProName.text())
        #评估产品
        if self.exInputOption.currentText() != ' ':
            self.exTechRevProduct.setText(self.exInputOption.currentText())
        #评审时间
        if self.exTAtime.text() != ' ':
            self.exTechReTime.setText(self.exTAtime.text())
        #技术能力引用
        self.exTechPart.setCurrentIndex(self.exPart.currentIndex())
        self.exTechInjection.setCurrentIndex(self.exInjection.currentIndex())
        self.exTechMetal.setCurrentIndex(self.exMetal.currentIndex())
        self.exTechDecorate.setCurrentIndex(self.exDecorate.currentIndex())
        self.exTechWield.setCurrentIndex(self.exWield.currentIndex())

        #根据供应商名称索引填充省和市
        #连接数据库
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


        sql = "select province, city from Supplier where abbreviation like '{0}'".format(self.exTechSupName.text())
        #print(sql)
        cursor.execute(sql)
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
        rows = cursor.fetchall()
        print(rows)
        #nrows[0] = '省'，nrows[1] = '市'
        nrows = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)

        #赋值
        for row in nrows:
            self.exTechProvin.setText(row[0])
            self.exTechCity.setText(row[1])
         #断开连接
        cursor.close()
        db.close()

    #返回按钮
    def back_to_external(self):
        self.stackedWidget.setCurrentIndex(3)

        #-----------------2019.7.27外协件信息录入界面按钮功能实现------------

        #外协件信息录入界面，导出excel功能
    def ex_export_excel(self):

        #复制到指定文件夹
        directory1 = QFileDialog.getExistingDirectory(self,
                                    "选取文件夹",
                                    "C:/")    

        backPath = directory1  +  '\\外协件前期技术管理工作计划表模板'  + '.xlsx'
        print(backPath)
        original = os.getcwd()  + r'\resourse\\' + '外协件前期技术管理工作计划表模板.xlsx'
        #复制并重命名
        shutil.copy(original, backPath)

        #填充表格
        #openpyxl连接表格
        try:
            book = openpyxl.load_workbook(backPath)
        except:
            QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #xlwt连接sheet,行列号起始数皆为0
        try:
            #ws = book.get_sheet_by_name("sheet1")
            ws = book.worksheets[0]
        except:
            QMessageBox.about(self, '错误', 'sheet定位失败！')

        #进行数据库索引
        #连接数据库
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        if self.exProName.text() != '':
            sql = "select productName, externalTime, mouldTime from External where projectName like '{0}'".format(self.exProName.text())
        #print(sql)
            cursor.execute(sql)
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
            rows = cursor.fetchall()
        #print(rows)
        #nrows[0] = '产品名称'，nrows[0] = 'TA评审时间，及外协件评审时间'，nrows[0] = '模具评审时间'
            nrows = []
            for row in rows:
                lRow = [int_it(each) for each in row]
                nrows.append(lRow)
            print(nrows)
        #填写表格，产品名称：9,2；TA评审日期：9,8；模具评审时间：9,11，行列都从1开始计数
            i = 9
            for row in nrows:
            #产品名称：9,2
                ws.cell(i, 2, row[0])
            #TA评审日期：9,8
                ws.cell(i, 8, row[1])
            #模具评审时间：9,11
                ws.cell(i, 11, row[2])
                i = i + 1
   
            book.save(filename = backPath)
            QMessageBox.about(self, '提示', '导出完成！')
         #断开连接
            cursor.close()
            db.close()
        else:
            QMessageBox.about(self, '提示', '项目名为空，请填写后再导出！')
        
        #根据技术能力，筛选布点供应商下拉框选项----------------------2019.8.10
    def exSup_option_byTech(self,box):
        #连接数据库
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        exList = []
        List = []
        ex1 = self.exPart.currentText()
        ex2 = self.exInjection.currentText()
        ex3 = self.exMetal.currentText()
        ex4 = self.exDecorate.currentText()
        ex5 = self.exWield.currentText()
        #获取关键字列表
        exList = [ex1, ex2, ex3, ex4, ex5]
        #print(exList)
        #去除空元素
        List = [i for i in exList if i != '']
        #print(List)
        
        sql = "select  abbreviation,  city, technologyOne, technologyTwo, technologyThree, \
                    technologyFour, technologyFive, technologySix, technologySeven, technologyEight, technologyNine,\
                     technologyTen from Supplier"
        cursor.execute(sql)
#转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
        rows = cursor.fetchall()
        #print(rows)
        nrows = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)

#根据tech交叉筛选下拉框选项
        abb = []
        city = []
        
        for row in nrows:
            #若包含数据关键字，则提取前两个元素，0为简称，1位市
            if set(List) < set(row):
                abb.append(row[0])
                city.append(row[1])
        #print(abb)
        #添加clear选项
        abb.insert(0,' ')
        city.insert(0,' ')
 #添加下拉选项,只添加exsup1,剩下三个复制exsup1即可，每次都初始化选项
        box.clear()
        self.exSup2.clear()
        self.exSup3.clear()
        self.exSup4.clear()

        self.exSup1Place.clear()
        self.exSup2Place.clear()
        self.exSup3Place.clear()
        self.exSup4Place.clear()
        
        box.addItems(abb)
        self.exSup2.addItems(abb)   
        self.exSup3.addItems(abb)  
        self.exSup4.addItems(abb)
            
        #city选项
        self.exSup1Place.addItems(city)
        self.exSup2Place.addItems(city)
        self.exSup3Place.addItems(city)
        self.exSup4Place.addItems(city)

        #断开连接
        cursor.close()
        db.close()
  


        #导入数据库按钮-2019.8.4
    def ex_ta_inputDb(self):
             #用pyodbc连接数据库
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        #读取数据库内原有数据，用于判断即将导入的数据是否重复，若重复则不导入
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，nrows,外协件技术评审时间为externalTime
        cursor.execute("select OEM, projectName, projectNumber, department, productName, externalTime, exTa, mouldTime, \
                    exPart, exInjection, exMetal, exDecorate, exWield, distributionSup1,\
                    distributionSup2,  distributionSup3, distributionSup4, exSeleSup, exSup1Place, exSup2Place,\
                    exSup3Place, exSup4Place, exSeleSupPlace, sup1Mak, sup2Mak, sup3Mak, sup4Mak, \
                    exSeleSupTaMark from External")
        
        #获取数据库内已有数据
        rows = cursor.fetchall()
        nrows = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)

        #获取外协件界面输入信息
        #按sql语句项目排序
        OEM = self.exOEM.currentText()
        projectName = self.exProName.text()
        projectNumber = self.exProNum.text()
        department = self.exDepart.currentText()
        productName = self.exInputOption.currentText()
        extTechTime = self.exTAtime.text()
        exTa = self.exTa.text()
        mouldTime = self.exMouldTime.text()
        exPart = self.exPart.currentText()
        exInjection = self.exInjection.currentText()
        exMetal = self.exMetal.currentText()
        exDecorate = self.exDecorate.currentText()
        exWield = self.exWield.currentText()
        distributionSup1 = self.exSup1.currentText()
        distributionSup2 = self.exSup2.currentText()
        distributionSup3 = self.exSup3.currentText()
        distributionSup4 = self.exSup4.currentText()
        exSeleSup = self.exSeleSup.currentText()
        exSup1Place = self.exSup1Place.currentText()
        exSup2Place = self.exSup2Place.currentText()
        exSup3Place = self.exSup3Place.currentText()
        exSup4Place = self.exSup4Place.currentText()
        exSeleSupPlace = self.exSeleSupPlace.currentText()

        #这些分数因分别由技术评估、供应商质量评估界面提供
        #sup1Mak = self.exSup1TaMark.text()
        #sup2Mak = self.exSup2TaMark.text()
        #sup3Mak = self.exSup3TaMark.text()
        #sup4Mak = self.exSup4TaMark.text()
        #exSeleSupTaMark = self.exSeleSupTaMark.text()

        value = [OEM, projectName, projectNumber, department, productName, extTechTime, exTa, mouldTime, \
                    exPart, exInjection, exMetal, exDecorate, exWield, distributionSup1,\
                    distributionSup2,  distributionSup3, distributionSup4, exSeleSup, exSup1Place, exSup2Place,\
                    exSup3Place, exSup4Place, exSeleSupPlace]
        #共28列
        insert_sql = """
            insert into External(OEM, projectName, projectNumber, department, productName, externalTime, exTa, mouldTime, \
                    exPart, exInjection, exMetal, exDecorate, exWield, distributionSup1,\
                    distributionSup2,  distributionSup3, distributionSup4, exSeleSup, exSup1Place, exSup2Place,\
                    exSup3Place, exSup4Place, exSeleSupPlace) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
        lValue = [int_it(each) for each in value]
                #print(lValue)
        if lValue not in nrows:
            cursor.execute(insert_sql, value)
            cursor.commit()

        QMessageBox.about(self, '提示', '导入完成！')
        #断开连接
        cursor.close()
        db.close()



        #添加附件按钮
    def ex_from_excel(self):
        self.exExcelPath = ''
        #清空选项
        self.exInputOption.clear()
        #选取文件
        fileName, fileType = QFileDialog.getOpenFileName(self,
                                    "选取文件",
                                    os.getcwd() + "\\resourse\\appendix",
                                    "All Files (*);;Text Files (*.txt)")
        
        self.exExcelPath = fileName
        if self.exExcelPath != '':
            try:
                book = xlrd.open_workbook(self.exExcelPath)
            except:
                QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            try:
                sheet = book.sheet_by_name("Sheet1")
            except:
                QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


            #项目号填写
            self.exProNum.setText(sheet.cell(3, 2).value)
            #项目名称填写
            self.exProName.setText(sheet.cell(4, 2).value)

             #OEM填写
            oemIndex = 0
            for i in range(self.exOEM.count()):
                if sheet.cell(2, 2).value ==  self.exOEM.itemText(i):
                    oemIndex = i
            if oemIndex != 0:
                self.exOEM.setCurrentText(sheet.cell(2, 2).value)
            else:
                QMessageBox.about(self, '提示', 'OEM命名与数据库中命名不一致！')

            #开发部门填写
            depIndex = 0
            for i in range(self.exDepart.count()):
                if sheet.cell(5, 2).value ==  self.exDepart.itemText(i):
                    depIndex = i
            if depIndex != 0:
                #QMessageBox.about(self, '提示', '字符为： ' + sheet.cell(5, 2).value)
                self.exDepart.setCurrentText(sheet.cell(5, 2).value)
            else:
                QMessageBox.about(self, '提示', '开发部门命名与数据库中命名不一致！')

            #--------------------初始化产品名称下拉框选项-----------------------
            #获取产品名称
            rowList = sheet.col_values(1)
            proList = ['']
            #从第9列开始，去除空值
            for i in range(8, len(rowList) ):
                if rowList[i] != '':
                    proList.append(rowList[i])
            
            #初始化评估产品下拉框选项
            
            self.exInputOption.addItems(proList)



            #提示导入完成
            QMessageBox.about(self, '提示', '导入完成！')
        else:
            QMessageBox.about(self, '提示', '未选择附件！')

        #-----------------------将表格中的数据填入程序界面----------------
    def ex_excel_to_ui(self):    
        #连接表格
        try:
            book = xlrd.open_workbook(self.exExcelPath)
        except:
            QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        try:
            sheet = book.sheet_by_name("Sheet1")
        except:
            QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


       

        #TA评审日期填写，位于表格中第八列 
        #获取产品名称列数值
        rowList = sheet.col_values(1)
        
        #注意若日期为空会报错，已解决，xlrd导入空单元格为''
        #获取产品行号
        if self.exInputOption.currentText() in rowList:
            rowNumber = rowList.index(self.exInputOption.currentText())
        #赋值,首先将excel中的日期转化为元组，(year, month, day, hour, minute, nearest_second)，0为时间戳mod，1900
        #print(sheet.cell(rowNumber, 12).value)
        if sheet.cell(rowNumber, 7).value != '':
            rowTaDate = xlrd.xldate_as_tuple(sheet.cell(rowNumber, 7).value, 0)
        #将元组变为日期形式，1900/1/1
            TaDate = str(rowTaDate[0])+ '/' + str(rowTaDate[1]) + '/' + str(rowTaDate[2]) 
        #赋值TA评审日期
            self.exTAtime.setText(TaDate)
        else:
            self.exTAtime.setText('')

        #模具工装评审日期填写，位于表格中第十一列
         #赋值,首先将excel中的日期转化为元组，(year, month, day, hour, minute, nearest_second)，0为时间戳mod，1900
        #print(sheet.cell(rowNumber, 10).value)
        if sheet.cell(rowNumber, 10).value != '':
            rowMoDate = xlrd.xldate_as_tuple(sheet.cell(rowNumber, 10).value, 0)
        #将元组变为日期形式，1900/1/1
            MoDate = str(rowMoDate[0])+ '/' + str(rowMoDate[1]) + '/' + str(rowMoDate[2]) 
        #赋值TA评审日期
            self.exMouldTime.setText(MoDate)
        else:
            self.exMouldTime.setText('')

        
        


    #项目号索引，待开发

    #技术评估按钮
    def ex_tech_review(self):
        #print('try')
        #引用外协件录入控件内容
        #供应商名称

        self.stackedWidget.setCurrentIndex(5)

    #供应商开发质量评估按钮
    def ex_sup_qua_review(self):
        #print('try')
        self.stackedWidget.setCurrentIndex(6)

        #引用内容
        #项目名称，评估产品，评审时间，技术能力引用
        #供应商名称
        self.exQuaSupName.setText(self.exSeleSup.currentText())
        #所在地省，市
        self.exQuaProvin.setText(self.exTechProvin.text())
        self.exQuaCity.setText(self.exTechCity.text())
        #项目名称
        if self.exProName.text() != ' ':
            self.exQuaProName.setText(self.exProName.text())
        #评估产品
        if self.exInputOption.currentText() != ' ':
            self.exQuaReProduct.setText(self.exInputOption.currentText())
        #项目号
        if self.exProNum.text() != ' ':
            self.exQuaProNum.setText(self.exProNum.text())
        #技术能力引用
        self.exQuaPart.setCurrentIndex(self.exPart.currentIndex())
        self.exQuaInjection.setCurrentIndex(self.exInjection.currentIndex())
        self.exQuaMetal.setCurrentIndex(self.exMetal.currentIndex())
        self.exQuaDecorat.setCurrentIndex(self.exDecorate.currentIndex())
        self.exQuaWield.setCurrentIndex(self.exWield.currentIndex())





        #--------------------------此函数实现界面切换--------------------------------------------------------------------------------------------------

    def changePage(self, index):
        item = self.treeWidget.currentItem()
        if item.text(0) == '供应商信息录入':
            self.stackedWidget.setCurrentIndex(1)
        elif item.text(0) == '供应商信息查询':
            self.stackedWidget.setCurrentIndex(2)
        elif item.text(0) == '外协件信息录入':
            self.stackedWidget.setCurrentIndex(3)
        elif item.text(0) == '外协件信息查询':
            self.stackedWidget.setCurrentIndex(7)
        elif item.text(0) == '选项变更':
            self.stackedWidget.setCurrentIndex(8)


        #数据库连接
    def creat_db(self):
        self.db = QSqlDatabase.addDatabase("QODBC")
        dsn = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
        self.db.setDatabaseName(dsn)
        self.db.open()
        if not self.db.open():
            QMessageBox.information(self, '错误', 'pyqt5数据库未打开',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
    #重构关闭事件，关闭窗口时断开连接
    def closeEvent(self, event):
        self.db.close()
        QSqlDatabase.removeDatabase('qt_sql_default_connection')
        #-----------------------------------------------------------------------------------------------------------
        #|------------------------------------------固定选项相关函数--------------------------------------------------|
        #-----------------------------------------------------------------------------------------------------------

    #--------------下拉框选项--------------------2019.7.14更新不同分类选项
    #2019.7.21更新外协件信息录入页下拉框选项
    def combo_option(self, box, keyWord ):
        #--------------------采用pyodbc连接数据库，为下拉框提供选项------------------------------------------
        try:
            stra = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(stra)
            cursorCombo = db.cursor()
        
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #科技能力和省采用ConstanOption表中选项，简称和市选用Supplier表
        if keyWord == 'technologicalCapability' or keyWord == 'province' or keyWord == 'injection' or \
        keyWord == 'decoration'  or keyWord == 'metalForming'  or keyWord == 'welding'  or keyWord == 'parts' or\
        keyWord == 'oem' or keyWord == 'depName' or keyWord == 'TAer' :
            sql = "select " + keyWord + " from ConstantOption"
        elif keyWord == 'city' or keyWord == 'abbreviation':
            sql = "select " + keyWord + " from Supplier"
        elif keyWord == 'projectName' or keyWord == 'projectNumber':
            sql = "select " + keyWord + " from External"
        
        #QMessageBox.information(self, '提示', sql,QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        cursorCombo.execute(sql)
        rows = cursorCombo.fetchall()
        nrows = []
        for row in rows:
            #无需int化，item只接受str类型
            #lRow = [int_it(each) for each in row]
            nrows.append(row)
        #去除重复项
        dNrows = []
        for i in nrows:
            if i not in dNrows :
                dNrows.append(i)
        #QMessageBox.information(self, '错误', dNrows,QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #将选项字符化，item只接受str类型
        #print(dNrows[0])
        
        



        for index, element in enumerate(dNrows):
            
            box.addItem(element[0])
            item = box.model().item(index, 0)
            item.setCheckState(QtCore.Qt.Unchecked)

        #设置自动补全
        #self.itemList = []
        #for i in range(len(dNrows)):
            #self.itemList.append(str(dNrows[i][0]))
            #print(str(dNrows[i][0]))

        
        #box.setCompleter(self.completer)
        #断开连接
        cursorCombo.close()
        db.close()

    #此函数获取下拉复选框选项文本用于输入数据库
    def getText(self, box):
        
        text1 = self.supTech1.getCheckItem()
        text2 = self.supTech3.getCheckItem()
        text3 = self.supTech3.getCheckItem()
        text4 = self.supTech4.getCheckItem()
        text5 = self.supTech5.getCheckItem()
        #将所有下拉框选项放入一个列表，便于显示和输入数据库
        self.techText = text1+text2 +text3 +text4 +text5
        #print(self.techText)
        #用于外协件信息录入，布点供应商下拉框选项
        #text01 = self.exPart.getCheckItem()
        #str01 = text01[0]
        #print(str01)
        #text02 = self.exInjection.getCheckItem()
        #text03 = self.exMetal.getCheckItem()
        #text04 = self.exDecorate.getCheckItem()
        #text05 = self.exWield.getCheckItem()
        #将所有下拉框选项放入一个列表，便于查询
        #self.exTechOption = []
        #self.exTechOption = text01+text02 +text03 +text04 +text05
        #print(self.exTechOption)


        #print(self.techText)
    #此函数用于更新定点供应商选项，其选项为四个布点供应商选项的内容,当布点供应商选项被选中时执行此函数更新定点供应商选项
    def select_sup_option(self,combox,word):
        if word == 'city':
            sList = []
            for i in range(combox.count()):
                sList.append(self.exSeleSupPlace.itemText(i))
            if combox.currentText() not in sList:
                self.exSeleSupPlace.addItem(combox.currentText())
        elif word =='sup':
            sList = []
            for i in range(combox.count()):
                sList.append(self.exSeleSup.itemText(i))
            if combox.currentText() not in sList:
                self.exSeleSup.addItem(combox.currentText())
        #-----------------------------------------------------------------------------------------------------------
        #|------------------------------------------按钮类相关函数--------------------------------------------------|
        #-----------------------------------------------------------------------------------------------------------

        #-------------------------单个导入-获取附件路径------------------------
    def get_path(self):
        fileName = ''
        fileName, fileType = QFileDialog.getOpenFileName(self,
                                    "选取文件",
                                    os.getcwd() + "\\resourse\\appendix",
                                    "All Files (*);;Text Files (*.txt)")
        if fileName != '':
            QMessageBox.about(self, '提示', '附件路径为： ' + fileName)
            #新增将选取的文件复制到appendix功能-2019.7.9
            #backPath = os.getcwd()  + r'\resourse\db_backup' '\\'+ 'supplier_' + nowtime + '.mdb'
            f = os.path.basename(fileName)
            #需要复制的文件全路径
            backPath = os.getcwd()  + r'\resourse\\appendix\\' + f
            #此变量用于储存路径到数据库内
            self.filePath = backPath
            #构建存储路径
            appendixPath = os.getcwd() + r'\resourse\\appendix\\'
            #判断文件是否存在
            for root, dirs, files in os.walk(appendixPath):
                if f not in files:
            #复制并重命名
                    shutil.copy(fileName, backPath)
            QMessageBox.about(self, '提示', '已获取路径！')

            #self.fileType = fileType
            #QMessageBox.information(self, '提示', self.filePath,QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        else:
            QMessageBox.about(self, '提示', '未选择附件！')
        #-------------------------批量输入展示------------------------
    def sup_more_view(self):
        #连接表格
        fileName = ''
        fileName, fileType = QFileDialog.getOpenFileName(self,
                                    "选取文件",
                                    os.getcwd() + "\\resourse\\appendix",
                                    "All Files (*);;Text Files (*.txt)")
        #批量导入
        #连接表格
        if fileName != '':
            try:
                book = xlrd.open_workbook(fileName)
            except:
                QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

            try:
                sheet = book.sheet_by_name("Sheet1")
            except:
                QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #设置model模型
            self.Mmodel = QStandardItemModel(sheet.nrows, 18)
            #设置表头
            self.Mmodel.setHorizontalHeaderLabels(['供应商名称', '简称','所在省', '所在市', '工艺能力1', '工艺能力2','工艺能力3', 
                '工艺能力4', '工艺能力5', '工艺能力6','工艺能力7','工艺能力8','工艺能力9','工艺能力10', '技术能力', '模具能力', 
                '公司简介','公司详情'])
            #表格内除去第一行表头
            for row in range(1, sheet.nrows):
                for column in range(19):
                    item = QStandardItem(" %s" %sheet.cell(row,column).value)
                    self.Mmodel.setItem(row-1, column, item)
            #连接view和model
            self.subInputView.setModel(self.Mmodel)
            # 自适应窗口
            self.subInputView.horizontalHeader().setStretchLastSection(True)
            self.subInputView.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        else:
            QMessageBox.about(self, '提示', '未选择文件！')

        #-------------------------单个输入展示------------------------
    def one_input_view(self):
        #设置model模型
        self.Omodel = QStandardItemModel(1, 19)
        #设置表头
        self.Omodel.setHorizontalHeaderLabels(['供应商名称', '简称','所在省', '所在市', '工艺能力1', '工艺能力2','工艺能力3', 
            '工艺能力4', '工艺能力5', '工艺能力6','工艺能力7','工艺能力8','工艺能力9','工艺能力10','设备能力', '技术能力', 
            '模具能力', '公司简介','公司详情'])
        #获取文本框内容
        supplierName = self.supName.text()
        abbreviation = self.supAbbreviation.text()
        province = self.supProvince.text()
        city = self.supCity.text()
        if self.techText != ' ':
            while len(self.techText) < 11:
                self.techText.append('')

            technology1 = self.techText[0]
            technology2 = self.techText[1]
            technology3 = self.techText[2]
            technology4 = self.techText[3]
            technology5 = self.techText[4]
            technology6 = self.techText[5]
            technology7 = self.techText[6]
            technology8 = self.techText[7]
            technology9 = self.techText[8]
            technology10 = self.techText[9]
               
            #technologyOne = self.text1
            #technologyTwo = self.supTech2.currentText()
            #technologyThree = self.supTech3.currentText()
            #technologyFour = self.supTech4.currentText()
            #technologyFive = self.supTech5.currentText()
            equipment = self.supEquip.text()
            technology = self.supTechRearch.currentText()
            model = self.supModelRearch.currentText()
            comBreif = self.supBrief.toPlainText()
            #采用qdialog获取文本信息，comIntroduction
            comIntroduction = self.filePath


            self.value = [supplierName, abbreviation, province, city, technology1, technology2, technology3, technology4, 
            technology5, technology6, technology7, technology8, technology9, technology10, equipment, technology, 
            model, comBreif, comIntroduction]
            #向表格输入内容
            for row in range(1):
                for column in range(19):
                    item = QStandardItem(" %s" %self.value[column])
                    self.Omodel.setItem(row, column, item)
            #连接view和model
            self.subInputView.setModel(self.Omodel)

        #-------------------------导入数据------------------------
    def input_into_db(self):
        #批量导入
        #连接表格
        try:
            book = xlrd.open_workbook('resourse\\供应商信息批量导入模板.xlsx')
        except:
            QMessageBox.information(self, '错误', '表格连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        try:
            sheet = book.sheet_by_name("Sheet1")
        except:
            QMessageBox.information(self, '错误', 'sheet定位失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #用pyodbc连接数据库
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        #print('connection success!')
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，nrows
        cursor.execute("select supplierName, abbreviation, province, city, technologyOne, technologyTwo, technologyThree, \
                    technologyFour, technologyFive, technologySix, technologySeven, technologyEight, technologyNine,\
                     technologyTen,  equipment, technology, model, comBreif, comIntroduction from Supplier")
        
        #获取数据库内已有数据，用于与即将导入的批量数据进行对比，不重复的才进行导入
        rows = cursor.fetchall()
        nrows = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)
        #print(nrows)
        #如果为批量导入，self.value储存单个导入的值，若为空则not self.value为true
        if not self.value:
            #改动时注意缩进，批量输入方式为判断一行输入一行
            for i in range(1,sheet.nrows):
                supplierName = sheet.cell(i,0).value
                abbreviation = sheet.cell(i,1).value
                province = sheet.cell(i,2).value
                city = sheet.cell(i,3).value
                technology1 = sheet.cell(i,4).value
                technology2 = sheet.cell(i,5).value
                technology3 = sheet.cell(i,6).value
                technology4 = sheet.cell(i,7).value
                technology5 = sheet.cell(i,8).value
                technology6 = sheet.cell(i,9).value
                technology7 = sheet.cell(i,10).value
                technology8 = sheet.cell(i,11).value
                technology9 = sheet.cell(i,12).value
                technology10 = sheet.cell(i,13).value
                equipment = sheet.cell(i,14).value
                technology = sheet.cell(i,15).value
                model = sheet.cell(i,16).value
                comBreif = sheet.cell(i,17).value
                comIntroductionMore = sheet.cell(i,18).value
                #复制文件到appendix
                if comIntroductionMore != '':
                    #获取文件名
                    f = os.path.basename(comIntroductionMore)
                    #构建储存路径
                    backPath = os.getcwd()  + r'\resourse\\appendix\\' + f
                    #判断文件夹中是否存在文件
                    appendixPath = os.getcwd() + r'\resourse\\appendix\\'
                    for root, dirs, files in os.walk(appendixPath):
                        if f not in files:
                #复制并重命名
                            shutil.copy(comIntroductionMore, backPath)

                value = [supplierName, abbreviation, province, city, technology1, technology2, technology3, technology4, 
                technology5, technology6, technology7, technology8, technology9, technology10, equipment, technology, model,
                 comBreif, comIntroductionMore]
    #print(value)
                insert_sql = """
            insert into Supplier(supplierName, abbreviation, province, city, technologyOne, technologyTwo, technologyThree, \
                    technologyFour, technologyFive, technologySix, technologySeven, technologyEight, technologyNine,\
                     technologyTen, equipment, technology, model, comBreif, comIntroduction) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
                lValue = [int_it(each) for each in value]
                #print(lValue)
                if lValue not in nrows:
                    cursor.execute(insert_sql, value)
                    cursor.commit()

        #若为单个导入
        else:
            insert_sql = """
            insert into Supplier(supplierName, abbreviation, province, city, technologyOne, technologyTwo, technologyThree, \
                    technologyFour, technologyFive, technologySix, technologySeven, technologyEight, technologyNine,\
                     technologyTen, equipment, technology, model, comBreif, comIntroduction) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            OValue = [int_it(each) for each in self.value]
            if OValue not in nrows:
                cursor.execute(insert_sql, self.value)
                cursor.commit()

        QMessageBox.about(self, '提示', '导入完成！')
        #断开连接
        cursor.close()
        db.close()
        #清楚旧选项，添加新选项（最优解位判断选项，新选项就加入box=========待完成）
        self.supQcity.clear()
        self.supQbrief.clear()
        self.combo_option(self.supQcity, 'city')
        self.combo_option(self.supQbrief, 'abbreviation')

        #-------------------------清空数据------------------------

    def clear_all(self):
        #清空文本框
        self.supName.clear()
        self.supAbbreviation.clear()
        self.supProvince.clear()
        self.supCity.clear()
        self.supEquip.clear()
        #self.subMould.clear()
        #self.subTech.clear()
        #清空text
        self.supBrief.clear()
        #通过设置空值清空下拉框,利用setCurrentIndex(int index)方法设置空值
        #供应商信息录入页（2）下拉框
        self.supTech1.setCurrentIndex(-1)
        self.supTech2.setCurrentIndex(-1)
        self.supTech3.setCurrentIndex(-1)
        self.supTech4.setCurrentIndex(-1)
        self.supTech5.setCurrentIndex(-1)
        self.supTechRearch.setCurrentIndex(-1)
        self.supModelRearch.setCurrentIndex(-1)
        #供应商信息查询页（3）下拉框
        self.supQtech.setCurrentIndex(-1)
        self.supQbrief.setCurrentIndex(-1)
        self.supQprovin.setCurrentIndex(-1)
        self.supQcity.setCurrentIndex(-1)

        #通过设置空model清空view
        Omodel = QStandardItemModel()
        self.subInputView.setModel(Omodel)
        #清空查询tablewidget
        self.supQview.clearContents()

        #清空外协件信息录入页
        #下拉框
        self.exOEM.setCurrentIndex(-1)
        self.exDepart.setCurrentIndex(-1)
        self.exInputOption.setCurrentIndex(-1)
        self.exPart.setCurrentIndex(-1)
        self.exInjection.setCurrentIndex(-1)
        self.exMetal.setCurrentIndex(-1)
        self.exDecorate.setCurrentIndex(-1)
        self.exWield.setCurrentIndex(-1)
        self.exSup1.setCurrentIndex(-1)
        self.exSup2.setCurrentIndex(-1)
        self.exSup3.setCurrentIndex(-1)
        self.exSup4.setCurrentIndex(-1)
        self.exSeleSup.setCurrentIndex(-1)
        self.exSup1Place.setCurrentIndex(-1)
        self.exSup2Place.setCurrentIndex(-1)
        self.exSup3Place.setCurrentIndex(-1)
        self.exSup4Place.setCurrentIndex(-1)
        self.exSeleSupPlace.setCurrentIndex(-1)
        #文本框
        self.exProNumLine.clear()
        self.exProName.clear()
        self.exProNum.clear()
        self.exTAtime.clear()
        self.exTa.clear()
        self.exMouldTime.clear()
        self.exSup1TaMark.clear()
        self.exSup2TaMark.clear()
        self.exSup3TaMark.clear()
        self.exSup4TaMark.clear()
        self.exSeleSupTaMark.clear()

        #技术评估界面，index = 5
        #Qlineedit
        self.exTechSupName.clear()
        self.exTechProvin.clear()
        self.exTechCity.clear()
        self.exTechProName.clear()
        self.exTechReTime.clear()
        #分数
        self.exTechM1.clear()
        self.exTechM2.clear()
        self.exTechM3.clear()
        self.exTechM4.clear()
        self.exTechM5.clear()
        self.exTechM6.clear()
        self.exTechM7.clear()
        self.exTechM8.clear()
        self.exTechM9.clear()
        self.exTechM10.clear()
        self.exTechM11.clear()
        self.exTechM12.clear()
        self.exTechM13.clear()
        self.exTechTotalMark.clear()

        #Qtextedit
        self.exTechRevProduct.clear()
        self.exTechDesF1.clear()
        self.exTechDesF2.clear()
        self.exTechDesF3.clear()
        self.exTechDesF4.clear()
        self.exTechDesF5.clear()
        self.exTechDesF6.clear()
        self.exTechDesF7.clear()
        self.exTechDesF8.clear()
        self.exTechDesF9.clear()
        self.exTechDesF10.clear()
        self.exTechDesF11.clear()
        self.exTechDesF12.clear()
        self.exTechDesF13.clear()

        self.exTechDesA1.clear()
        self.exTechDesA2.clear()
        self.exTechDesA3.clear()
        self.exTechDesA4.clear()
        self.exTechDesA5.clear()
        self.exTechDesA6.clear()
        self.exTechDesA7.clear()
        self.exTechDesA8.clear()
        self.exTechDesA9.clear()
        self.exTechDesA10.clear()
        self.exTechDesA11.clear()
        self.exTechDesA12.clear()
        self.exTechDesA13.clear()

        self.exTechPart.setCurrentIndex(-1)
        self.exTechInjection.setCurrentIndex(-1)
        self.exTechMetal.setCurrentIndex(-1)
        self.exTechDecorate.setCurrentIndex(-1)
        self.exTechWield.setCurrentIndex(-1)
        self.exTechSupNum.setCurrentIndex(-1)

        #供应商质量评估界面-----------------------
        #Qlineeidt
        self.exQuaSupName.clear()
        self.exQuaProvin.clear()
        self.exQuaCity.clear()
        self.exQuaProName.clear()
        self.exQuaProNum.clear()

        self.exQuaM1.clear()
        self.exQuaM2.clear()
        self.exQuaM3.clear()
        self.exQuaM4.clear()
        self.exQuaM5.clear()
        self.exQuaM6.clear()
        self.exQuaM7.clear()
        self.exQuaM8.clear()

        self.exQuaReProduct.clear()
        #下拉框
        self.exQuaPart.setCurrentIndex(-1)
        self.exQuaInjection.setCurrentIndex(-1)
        self.exQuaMetal.setCurrentIndex(-1)
        self.exQuaDecorat.setCurrentIndex(-1)
        self.exQuaWield.setCurrentIndex(-1)
        
        #第八页------------------------------------外协件信息查询
        self.exQQEMcombo.setCurrentIndex(-1)
        self.exQProNameCombo.setCurrentIndex(-1)
        self.exQProNumCombo.setCurrentIndex(-1)
        self.exQTAcombo.setCurrentIndex(-1)
        self.exQInjection.setCurrentIndex(-1)
        self.exQDecorate.setCurrentIndex(-1)
        self.exQPart.setCurrentIndex(-1)
        self.exQWield.setCurrentIndex(-1)
        self.exQMetal.setCurrentIndex(-1)



        #-------------------------备份数据库------------------------
    def back_up_db(self):

        #备份名称采用name+time的形式
        nowtime=time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))
        backPath = os.getcwd()  + r'\resourse\db_backup' '\\'+ 'supplier_' + nowtime + '.mdb'

        original = os.getcwd()  + r'\resourse\\' + 'supplier.mdb'
        #复制并重命名
        shutil.copy(original, backPath)
        QMessageBox.about(self, '提示', '备份完成！')

        #-------------------------查询------------------------
    def db_query(self):

        #要实现任意关键字交叉查询，需要遍历所有关键字组合情况，四个关键字共有16中情况
        #不同情况改变部分SQL语句，采用sql02表示

        #获取关键字文本
        techText = self.supQtech.currentText()
        breifText = self.supQbrief.currentText()
        provinceText = self.supQprovin.currentText()
        cityText = self.supQcity.currentText()

        #含有一个关键字，共四种情况
        #1.1简称查询
        if techText == ''  and breifText !='' and provinceText == '' and cityText == '':
            sql02 = "where abbreviation like '{0}'".format(breifText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #1.2技术能力查询
        if techText != ''  and breifText =='' and provinceText == '' and cityText == '':
            sql02 = "where technologyOne like '{0}' or technologyTwo like '{0}'\
            or technologyThree like '{0}' or technologyFour like '{0}' or technologyFive like '{0}'".format(techText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #1.3省查询
        if techText == ''  and breifText =='' and provinceText != '' and cityText == '':
            sql02 = "where province like '{0}'".format(provinceText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #1.4市查询
        if techText == ''  and breifText =='' and provinceText == '' and cityText != '':
            sql02 = "where city like '{0}'".format(cityText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)

        #两个关键字，共六种情况
        #2.1 技术能力+简称
        if techText != ''  and breifText !='' and provinceText == '' and cityText == '':
            sql02 = "where abbreviation like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')".format(breifText,techText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #2.2 技术能力+省
        if techText != ''  and breifText =='' and provinceText != '' and cityText == '':
            sql02 = "where province like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')".format(provinceText,techText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #2.3技术能力+市
        if techText != ''  and breifText =='' and provinceText == '' and cityText != '':
            sql02 = "where city like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')".format(cityText,techText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #2.4 简称+省
        if techText == ''  and breifText !='' and provinceText != '' and cityText == '':
            sql02 = "where abbreviation like '{0}' and ( province like '{1}')".format(breifText,provinceText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #2.5 简称+市
        if techText == ''  and breifText !='' and provinceText == '' and cityText != '':
            sql02 = "where abbreviation like '{0}' and ( city like '{1}')".format(breifText,cityText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #2.6 市+省
        if techText == ''  and breifText =='' and provinceText != '' and cityText != '':
            sql02 = "where city like '{0}' and ( province like '{1}')".format(cityText,provinceText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)

        #三个关键字，共四种情况
        #3.1 技术+简称+省
        if techText != ''  and breifText !='' and provinceText != '' and cityText == '':
            sql02 = "where abbreviation like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')\
            and ( province like '{2}')".format(breifText,techText,provinceText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #3.2 技术+简称+市
        if techText != ''  and breifText !='' and provinceText == '' and cityText != '':
            sql02 = "where abbreviation like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')\
            and ( city like '{2}')".format(breifText,techText,cityText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #3.3 技术+省+市
        if techText != ''  and breifText == '' and provinceText != '' and cityText != '':
            sql02 = "where city like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')\
            and ( province like '{2}')".format(cityText,techText,provinceText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)
        #3.4 简称+省+市
        if techText == ''  and breifText != '' and provinceText != '' and cityText != '':
            sql02 = "where city like '{0}' and ( abbreviation like '{1}')\
            and ( province like '{2}')".format(cityText,breifText,provinceText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)

        #四个关键字
        if techText != ''  and breifText !='' and provinceText != '' and cityText != '':
            sql02 = "where abbreviation like '{0}' and ( technologyOne like '{1}' or technologyTwo like '{1}'\
            or technologyThree like '{1}' or technologyFour like '{1}' or technologyFive like '{1}')\
            and ( province like '{2}') and (city like '{3}')".format(breifText,techText,provinceText,cityText)
            #QMessageBox.information(self, '提示', sql02, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) 
            self.execute_sql(sql02)


    def execute_sql(self,sql02):
        try:
            str = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
            db = pyodbc.connect(str)
            cursor = db.cursor()
        except:
            QMessageBox.information(self, '错误', 'pyodbc数据库连接失败！',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        sql = " select supplierName,abbreviation, province, city, technologyOne, technologyTwo, technologyThree, \
                    technologyFour, technologyFive, equipment, technology, model, comBreif, comIntroduction  from Supplier " + sql02
        #print(sql)
        #QMessageBox.information(self, '提示', sql, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        cursor.execute(sql)
        #QMessageBox.information(self, '提示', self.supQbrief.currentText()+ 
                        #self.supQprovin.currentText()+ self.supQcity.currentText()+ self.supQtech.currentText()
                       # , QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
       
        #转化格式，将数据库中已有的行数据格式转化为从excel中读取的数据格式，当数据既有字符串又有数字时需要转化，nrows          
        rows = cursor.fetchall()
        nrows = []
        path = []
        for row in rows:
            lRow = [int_it(each) for each in row]
            nrows.append(lRow)

        #获取查询结果中的地址放入path
        for i in range(len(nrows)):
            rowSlice = nrows[i]
            #地址不为空则放入path
            if rowSlice[13] !='':
                path.append(rowSlice[13])
        #设置行列
        self.supQview.setColumnCount(14)
        self.supQview.setRowCount(len(nrows))
        self.supQview.setHorizontalHeaderLabels(['供应商名称', '简称','所在省', '所在市', '工艺能力1', '工艺能力2','工艺能力3', 
            '工艺能力4', '工艺能力5', '设备能力', '技术能力', '模具能力', '公司简介','公司详情']) 
        #不可编辑 
        self.supQview.setEditTriggers(QTableView.NoEditTriggers)
        #自适应窗口大小
        self.supQview.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        #设置查看按钮
        #复制path用于创建以地址命名的按钮，按钮命名则以path[]为来源
        sbtn = path[:]
        #此变量用于创建不同button
        number = 0
        for row in range(len(nrows)):
            #添加控件
            if rowSlice[13] != '':
                    #批量添加控件的方式，目前要返回不同数据，只能靠以数据命名不同按钮的方式来获得
                    #构造Qbutton实例
                    sbtn[number] = QPushButton()
                    #以数据源命名 
                    sbtn[number].setText(path[number])
                    #单击按钮时返回按钮名称，即数据源 
                    sbtn[number].clicked.connect(lambda: self.open_file(self.sender().text()))
                    #button格式
                    sbtn[number].setStyleSheet("QPushButton{margin:3px};")
                    #添加到第13列
                    self.supQview.setCellWidget(row, 13, sbtn[number])
                    number += 1

            #添加查询结果
            for column in range(13):
                #rowslice获取nrows中一行的数据，与获取path的rowSlice不同
                rowslice = nrows[row]
                item = QTableWidgetItem(" %s" % rowslice[column])
                self.supQview.setItem(row, column, item)
        cursor.close() 
        db.close()  

       
     #-------------------------打开指定路径文件------------------------采用os.system()，需判断后缀名          
    def open_file(self,file):
            #获取路径与扩展名，用于构建打开路径
            dirname,extName=os.path.split(file)
            #路径构建
            file_path = os.getcwd() + "\\resourse\\appendix\\" +extName
            #os.system(),subprocess.Popen(),os.startfile()这几种方式都行,最后一个简单粗暴
            os.startfile(file_path)

class ExecDatabaseDemo(Ui_Form):

    def __init__(self, parent=None):
        super(ExecDatabaseDemo , self).__init__(parent)
        
        self.db = QSqlDatabase.addDatabase("QODBC")
        dsn = "DRIVER=Microsoft Access Driver (*.mdb, *.accdb);FIL={MS Access};DBQ=" + os.getcwd() + "\\resourse\\supplier.mdb"
        self.db.setDatabaseName(dsn)
        self.db.open()
        if not self.db.open():
            QMessageBox.information(self, '错误', 'pyqt5数据库未打开',QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
         

    def closeEvent(self, event):
        # 关闭数据库
        QSqlDatabase.removeDatabase('qt_sql_default_connection')
        self.db.close()
        
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    Form = QtWidgets.QWidget()
    demo = ExecDatabaseDemo()

    demo.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())        


