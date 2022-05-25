from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QDate
from main_ui import Ui_Gotin
from main_th import WorkThread
import sys, os

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_Gotin()
        self.ui.setupUi(self)
        self.init()  # 初始化自定义
        self.bind()  # 信号槽绑定调用

    # 信号槽绑定函数
    def bind(self):
        self.ui.exitBtn.clicked.connect(lambda :app.quit())
        self.ui.newsBtn.clicked.connect(self.newsBtnClick)
        self.ui.tempBtn.clicked.connect(self.tempBtnClick)
        self.ui.resetBtn.clicked.connect(self.init)
        self.ui.calcBtn.clicked.connect(self.calcBtnClick)

    # 初始化函数
    def init(self):
        self.newsPath = ''  # 消息文件路径
        self.tempPath = ''  # 模板文件路径
        self.ui.dateStart.setDate(QDate.currentDate())  # 设置当前日期
        self.ui.dateEnd.setDate(QDate.currentDate())
        self.ui.progressBar.setValue(0)
        self.ui.newsPath.setText('尚未打开消息文本文件')
        self.ui.tempPath.setText('尚未打开 Excel 表文件')
        MainWindow.setFixedSize(self, 800, 245)

    # 消息文件打开按钮点击事件
    def newsBtnClick(self):
        self.newsPath = QFileDialog.getOpenFileName(self, '选择文件', '.', '文本文件 | (*.txt)')[0]
        if self.newsPath:
            self.ui.newsPath.setText(self.newsPath[0:16] + '.../' + os.path.basename(self.newsPath))

    # 模板文件打开按钮点击事件
    def tempBtnClick(self):
        self.tempPath = QFileDialog.getOpenFileName(self, '选择文件', '.', 'Excel文件 | (*.xlsx)')[0]
        if self.tempPath:
            self.ui.tempPath.setText(self.tempPath[0:16] + '.../' + os.path.basename(self.tempPath))

    # 计算导出按钮点击事件
    def calcBtnClick(self):
        if not self.newsPath or not self.tempPath:
            QMessageBox.warning(self, '警告', '未打开文件!')
            return
        if self.ui.dateStart.date() > self.ui.dateEnd.date():
            tempDate = self.ui.dateStart.date()
            self.ui.dateStart.setDate(self.ui.dateEnd.date())
            self.ui.dateEnd.setDate(tempDate)
        with open(self.newsPath, 'r', encoding='utf-8') as f:
            content = f.read()
        self.ui.progressBar.setValue(0)
        self.ui.calcBtn.setEnabled(False)
        self.ui.calcBtn.setText('计算中')
        # 线程开启
        self.thread = WorkThread(content, self.tempPath, self.ui.dateStart.date(), self.ui.dateEnd.date())
        self.thread.endSignal.connect(self.threadEnd)   # 绑定线程结束信号
        self.thread.progressSignal.connect(self.addProgress)    # 绑定进度条信号
        self.thread.start()

    # 线程结束信号
    def threadEnd(self):
        self.ui.progressBar.setValue(100)
        QMessageBox.information(self, '提示', '成功导出!')
        self.ui.calcBtn.setEnabled(True)
        self.ui.calcBtn.setText('计算导出')

    # 进度条信号
    def addProgress(self, prog):
        currentProgress = self.ui.progressBar.value()
        currentProgress += round(100 / prog)
        self.ui.progressBar.setValue(currentProgress)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    root = MainWindow()
    root.show()
    sys.exit(app.exec_())

