'''
created by Highwindmx @ 20190711
For Batch Report Searching and Downloading
'''
import os
import sys 
from PyQt5.QtWidgets import (QApplication, QMainWindow,\
                            QWidget,QPushButton,QPlainTextEdit,QTableWidget,QTableWidgetItem,\
                            QGridLayout)
from PyQt5.QtGui import QIcon,QColor
from PyQt5.QtCore import Qt
import pandas as pd
import shutil

######################### 使用时可根据情况修改 ###################
SPLITMARK = " | "
DESTDIR = "D:/test" # 报告默认下载路径，后期可以用FileDialog来实现
DATABASE = os.path.abspath("./res/database/index.xlsx") #索引表格路径
######################### 使用时可根据情况修改 ###################

class MainWindow(QMainWindow): 
    def __init__(self, parent=None): 
        super().__init__()
        self.initUI()
        self.index = pd.read_excel(DATABASE)

    def initUI(self):
        central_QW = QWidget(self)
        self.setCentralWidget(central_QW)
        self.layoutMainGrid = QGridLayout() 
        central_QW.setLayout(self.layoutMainGrid)

        self.input_list = QPlainTextEdit("请输入终端客户编码",self)
        self.input_list.setMaximumWidth(300)

        self.search_btn = QPushButton('点击开始查找', self)

        self.output_list = QTableWidget(0,3,self)
        self.output_list.setHorizontalHeaderLabels(["终端客户编码","文件数量","路径"])

        self.download_btn = QPushButton('下载', self)
        self.download_btn.setToolTip('点击下载导出')

        self.status_bar = self.statusBar()

        self.layoutMainGrid.addWidget(self.input_list,   0,0,1,1)
        self.layoutMainGrid.addWidget(self.output_list,  0,1,1,1)
        self.layoutMainGrid.addWidget(self.search_btn,   1,0,1,1)
        self.layoutMainGrid.addWidget(self.download_btn, 1,1,1,1)
        self.setGeometry(200, 200, 1000, 600)
        self.setWindowTitle("GenetronHealth Report Searcher")
        self.setWindowIcon(QIcon("./res/icon/sunme.ico"))
        self.show()

        self.input_list.textChanged.connect(self.showIDnums)
        self.search_btn.clicked.connect(self.searchReports)
        self.download_btn.clicked.connect(self.downloadReports)

    def showIDnums(self):
        input_doc = self.input_list.toPlainText()
        lines = input_doc.split("\n")
        self.output_list.setRowCount(len(lines)) 
        i = 0
        for l in lines:
            n_item = QTableWidgetItem(l.strip())
            self.output_list.setItem(i,0,n_item)
            i+=1
        self.status_bar.showMessage(f"已输入 {len(lines)} 条待查询")

    def searchReports(self):
        total_num = 0
        for i in range(0,self.output_list.rowCount()):
            pid = self.output_list.item(i,0).text()
            #path = self.index.loc[self.index["PID"].str.contains(pid), "Path"]
            path = self.index.loc[self.index["PID"].isin([pid]), "Path"]
            path_num = len(list(path))
            total_num += path_num
            self.output_list.setItem(i,1, QTableWidgetItem(str(path_num)))
            self.output_list.setItem(i,2, QTableWidgetItem(SPLITMARK.join(list(path))))
            self.output_list.resizeColumnsToContents()
            if path_num > 1:
                self.output_list.item(i,1).setBackground(QColor(0,0,255))
            if path_num < 1:
                self.output_list.item(i,1).setBackground(QColor(255,0,0))
        self.status_bar.showMessage(f"共找到 {total_num} 条报告记录")
    
    def downloadReports(self):
        sn = 0
        fn = 0
        log = ""
        for i in range(0,self.output_list.rowCount()):
            path = self.output_list.item(i,2).text().split(SPLITMARK)
            for p in path:
                try:
                    shutil.copy(p, DESTDIR)
                except Exception as e:
                    fn+=1
                    log+=f"拷贝失败：共计{fn}，{os.path.basename(p)} 错误为：{e}\n"
                else:
                    sn+=1
                    log+=f"拷贝成功：共计{sn}，已完成：{os.path.basename(p)}\n"
        with open(os.path.join(DESTDIR,"log.txt"), "w") as f:
            f.write(log)
            
app = QApplication(sys.argv) 
w = MainWindow() 
sys.exit(app.exec())