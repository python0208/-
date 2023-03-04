import time
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib as mpl

from PySide2.QtWidgets import QApplication, QTableWidgetItem, QMessageBox
from PySide2.QtUiTools import QUiLoader
import sys


class Baidui_tr:
    def __init__(self):  # 界面导入 和事件响应
        """
        --------------------------------------------------------------------
        实例化界面对象，使用pyside2的QUiLoader方法获取ui界面
        绑定按钮事件在指定函数 状态栏内容展示
        """
        self.updat_window = QUiLoader().load('书籍修改.ui')


        self.statistic_window = QUiLoader().load('类别.ui')
        self.out_window = QUiLoader().load('退出.ui')
        self.out_window.pushButton.clicked.connect(self.out)
        self.out_window.pushButton_2.clicked.connect(self.close_out_window)

        self.delet_window = QUiLoader().load('删除信息.ui')
        self.delet_window.pushButton.clicked.connect(self.delete)
        # self.delet_window.pushButton.clicked.connect(self.delet_window.textBrowser.clear)
        self.delet_window.pushButton_2.clicked.connect(self.close_delete)

        self.main_window = QUiLoader().load('个人图书管理系统.ui')
        self.main_window.pushButton_7.clicked.connect(self.updat_window.show)
        self.updat_window.pushButton_2.clicked.connect(self.updat_window.close)
        self.updat_window.pushButton.clicked.connect(self.sure_updata)
        self.table = self.main_window.tableWidget
        self.statistic_window.pushButton.clicked.connect(self.main_window.textBrowser_3.clear)
        self.statistic_window.pushButton.clicked.connect(self.statistic_window.tableWidget_2.clear)

        self.main_window.pushButton_11.clicked.connect(self.statistic_window.show)
        self.statistic_window.pushButton.clicked.connect(self.statistic_window.close)

        self.main_window.pushButton_10.clicked.connect(self.out_windows)
        self.main_window.pushButton_11.clicked.connect(self.statistics)
        # self.main_window.pushButton_11.clicked.connect(self.main_window.textBrowser_3.clear)
        self.main_window.pushButton_5.clicked.connect(self.infoss)
        self.main_window.pushButton_6.clicked.connect(self.main_window.textBrowser_2.clear)
        self.main_window.pushButton_8.clicked.connect(self.matplt)
        self.main_window.pushButton_3.clicked.connect(self.search)
        self.main_window.pushButton_4.clicked.connect(self.delete_window)
        self.main_window.pushButton_9.clicked.connect(self.insert)
        self.main_window.pushButton.clicked.connect(self.save_info)
        self.main_window.pushButton_7.clicked.connect(self.updata_book)
        # self.main_windows = QUiLoader().load('借书.ui')
        # self.main_windows.pushButton.clicked.connect(self.close_borrow)
        # self.main_windows.pushButton_2.clicked.connect(self.borrow_book)
        # self.main_windows.pushButton_3.clicked.connect(self.returns)

        self.main_window.statusbar.showMessage('作者联系方式: 3181456558')

    def updata_book(self):
        """
        修改书籍内容
        """
        try:
            book_name = self.books()
            book_id = self.main_window.lineEdit_10.text()
            book_user = self.main_window.lineEdit_11.text()
            book_out = self.main_window.lineEdit_12.text()
            price = self.main_window.lineEdit_13.text()
            book_category = self.main_window.lineEdit_14.text()
            delete_book_info = self.deleteinfo()
            self.updat_window.textBrowser.append(str(delete_book_info))
            self.updat_window.textBrowser_2.append(book_name)
            self.updat_window.textBrowser_3.append(book_id
                                                   )
            self.updat_window.textBrowser_4.append(book_user)
            self.updat_window.textBrowser_5.append(book_out)
            self.updat_window.textBrowser_6.append(price)
            self.updat_window.textBrowser_7.append(book_category)
        except Exception as e:
            with open('日志.csv',mode='a')as f:
                f.write(str(e))
                f.write('\n')


    def sure_updata(self):

        book_name = self.books()
        df = pd.read_excel("图书.xlsx")
        try:
            print(df['书名'])
            for book in df['书名']:
                if book_name == book:
                    idex = list(df['书名']).index(book)
                    df = df.drop(idex)
                    df.to_excel('图书.xlsx', encoding='utf8', index=False)
            book_id = self.main_window.lineEdit_10.text()
            book_user = self.main_window.lineEdit_11.text()
            book_out = self.main_window.lineEdit_12.text()
            price = self.main_window.lineEdit_13.text()
            book_category = self.main_window.lineEdit_14.text()
            info = {'id': int(book_id), "书名": book_name, "作者": book_user, "价格": int(price), "出版单位": book_out, "类别": str(book_category)}
            df = pd.read_excel("图书.xlsx")
            df = df.append(pd.Series(info), ignore_index=True)
            df.to_excel('图书.xlsx', encoding='utf8', index=False)
            self.messagebox = QMessageBox()
            self.messagebox.information(self.messagebox, '欢迎使用', '保存成功')
        except Exception as e:
            with open('日志.csv', mode='a') as f:
                f.write(str(e))
                f.write('\n')

    def matplt(self):
        mpl.rcParams['font.sans-serif'] = ['SimHei']
        file_price = pd.read_excel('图书.xlsx')['类别']
        # file_x = pd.read_excel('图书.xlsx')['id']
        np_y = np.array(file_price)
        l = [i for i in range(len(file_price))]
        np_x = np.array(l)
        plt.subplot(122)
        plt.bar(np_y, np_x, color='red', label='sigmoid')
        plt.title('书籍分类')
        dic = {}
        for city in list(pd.read_excel('图书.xlsx')['出版单位']):
            if city in dic:
                dic[city] += 1
            else:
                dic[city] = 1
        c_x = [c for c in dic.keys()]
        c_s = [s for s in dic.values()]
        plt.subplot(121)
        plt.bar(c_x, c_s)
        plt.title('出版单位')

        plt.ioff()
        plt.show()

    def books(self):
        """
        看修改的书籍是否存在
        """
        books_ = pd.read_excel('图书.xlsx')
        find_book = list(books_['书名'])
        book = self.main_window.lineEdit_9.text()
        try:
            if book in find_book:
                return book
            else:
                return '不存在书籍'
        except Exception as e:
            with open('日志.csv', mode='a', encoding='utf8') as f:
                f.write(str(e))
                f.write('\n')

    def deleteinfo(self):
        """
        要修改书籍的基本信息

        """
        df = pd.read_excel('图书.xlsx')
        try:
            b = pd.read_excel('图书.xlsx')['书名']

            books = list(pd.read_excel('借阅.xlsx')['书名'])
            delet_book = self.books()
            for i in books:
                if delet_book == i:
                    dex = list(b).index(i)
                    info = df.loc[dex]
                    return info

            else:
                self.messagebox = QMessageBox()
                self.messagebox.information(self.messagebox, '不存在', '书籍不存在')

        except Exception as e:
           with open('日志.csv', mode='a') as f:
               f.write(e)
               f.write('\n')

    def infoss(self):
        """
        -------------------------------------------------------------------
        使用pandas读取excel文件，将读取的内容一一对应的放入QTableWidget控件中去
        每一个for循环表示QTableWidget中每一列所对应的内容
        QTableWidgetItem是QTableWidget控件可读取的对象

        """
        df = pd.read_excel('图书.xlsx')

        for i in range(1, len(df['书名']) + 1):
            id = (df['id'][i - 1])

            newitem_i = QTableWidgetItem(str(id))
            self.main_window.tableWidget.setItem(i, 0, newitem_i)
        for i in range(1, len(df['书名']) + 1):
            z = (df['书名'][i - 1])

            newitem_z = QTableWidgetItem(str(z))
            self.main_window.tableWidget.setItem(i, 1, newitem_z)
        for i in range(1, len(df['书名']) + 1):
            auth = df['作者'][i - 1]

            newitem_i = QTableWidgetItem(str(auth))
            self.main_window.tableWidget.setItem(i, 2, newitem_i)
        for i in range(1, len(df['书名']) + 1):
            price = (df['价格'][i - 1])

            newitem_i = QTableWidgetItem(str(price))
            self.main_window.tableWidget.setItem(i, 3, newitem_i)
        for i in range(1, len(df['书名']) + 1):
            fro = (df['出版单位'][i - 1])

            newitem_i = QTableWidgetItem(str(fro))
            self.main_window.tableWidget.setItem(i, 4, newitem_i)
        for i in range(1, len(df['书名']) + 1):
            time = (df['类别'][i - 1])

            newitem_i = QTableWidgetItem(str(time))
            self.main_window.tableWidget.setItem(i, 5, newitem_i)

    def search(self):
        """
        搜索功能
        -------------------------------------------------------------------
        实现通过ui界面中的lineEdit控件获取需要搜索的内容，点击对应的按钮转到该函数中
        对内容进行处理搜索
        搜索对象 遍历excel文件，按指定内容查找，获取符合的图书，最后返回到UI界面中的
        textBrowser控件中
        按类别 id 和作者也可以进行查找

        """
        df = pd.read_excel("图书.xlsx")
        name = self.main_window.lineEdit_7.text()
        if name in list(df['书名']):
            for i in df['书名']:
                if name in i:
                    dex = list(df['书名']).index(i)
                    info = df.loc[dex]
                    self.main_window.textBrowser_2.append(str(info))
        print(name
              )
        if (str(name)) in list(df['id']):
            for i in df['id']:
                if str(name) == str(i):
                    dex = list(df['id']).index(i)
                    info = df.loc[dex]
                    self.main_window.textBrowser_2.append(str(info))
        if name in list(df['作者']):
            for i in df['作者']:
                if str(name) in str(i):
                    dex = list(df['作者']).index(i)
                    info = df.loc[dex]
                    self.main_window.textBrowser_2.append(str(info))
        if name in list(df['类别']):
            for i in df['类别']:
                if str(name) in str(i):
                    dex = list(df['类别']).index(i)
                    info = df.loc[dex]
                    self.main_window.textBrowser_2.append(str(info))
        if name in list(df['出版单位']):
            for i in df['出版单位']:
                if str(name) in str(i):
                    dex = list(df['出版单位']).index(i)
                    info = df.loc[dex]
                    self.main_window.textBrowser_2.append(str(info))

    def delete_window(self):
        self.delet_window.show()
        book_name = self.main_window.lineEdit_8.text()

        df = pd.read_excel("图书.xlsx")
        for i in df['书名']:
            if book_name in i:
                dex = list(df['书名']).index(i)
                info = df.loc[dex]
                self.delet_window.textBrowser.append(str(info))

    def close_delete(self):
        self.delet_window.close()

    def delete(self):
        """
        删除功能
        -------------------------------------------------------------------
        从界面中输入需要删除的书籍，通过按钮连接到这个函数在excel中查找对应书籍
        使用try异常处理函数确保正常运行
        使用QMessageBox实现弹窗功能提示

        """

        book_name = self.main_window.lineEdit_8.text()
        df = pd.read_excel("图书.xlsx")
        try:
            print(df['书名'])
            for book in df['书名']:
                if book_name == book:
                    idex = list(df['书名']).index(book)
                    df = df.drop(idex)
                    df.to_excel('图书.xlsx', encoding='utf8', index=False)
                    self.messagebox = QMessageBox()
                    self.messagebox.information(self.messagebox, '删除', '删除成功')

        except Exception as e:
            self.messagebox = QMessageBox()
            self.messagebox.information(self.messagebox, '出现错误', str(e))

    def insert(self):
        """
        增加书籍
        --------------------------------------------------------------
        在界面上的添加书籍部分写入书籍内容，点击提交按钮把需要保存的书籍展示
        在QTableWidget控件中可以二次确让所保存的书籍信息是否有误
        self.main_window.lineEdit.text():获取文本框中所输入的内容
        self.main_window.tableWidget_2.setItem(0, 1, newitem_i):
        将输入的内容依次添加到对于的QTableWidget控件中展示

        :return: insert 需要保存到excel中的数据
        """

        try:
            i = self.main_window.lineEdit.text()
            name = self.main_window.lineEdit_2.text()
            author = self.main_window.lineEdit_3.text()
            price = self.main_window.lineEdit_4.text()
            unit = self.main_window.lineEdit_5.text()
            time = self.main_window.lineEdit_6.text()
            newitem_i = QTableWidgetItem(str(i))
            self.main_window.tableWidget_2.setItem(0, 1, newitem_i)
            newitem_i = QTableWidgetItem(str(name))
            self.main_window.tableWidget_2.setItem(1, 1, newitem_i)
            newitem_i = QTableWidgetItem(str(author))
            self.main_window.tableWidget_2.setItem(2, 1, newitem_i)
            newitem_i = QTableWidgetItem(str(price))
            self.main_window.tableWidget_2.setItem(3, 1, newitem_i)
            newitem_i = QTableWidgetItem(str(unit))
            self.main_window.tableWidget_2.setItem(4, 1, newitem_i)
            newitem_i = QTableWidgetItem(str(time))
            self.main_window.tableWidget_2.setItem(5, 1, newitem_i)
            inset = {'id': int(i), "书名": name, "作者": author, "价格": int(price), "出版单位": unit, "类别": str(time)}
            print(author)
            # df = df.append(inset, ignore_index=True)
            # df.to_excel('图书.xlsx', encoding='utf8', index=False)
            return inset
        except Exception as e:
            self.messagebox = QMessageBox()
            self.messagebox.information(self.messagebox, '出现错误', str(e))

    def save_info(self):
        """
        ------------------------------------------------------------------
        调用insert函数获取返回的数据
        使用pandas先读取原先的excel数据，再使用pandas中的to_excel方法将新添加的
        书籍保存

        """
        try:
            info = self.insert()
            df = pd.read_excel("图书.xlsx")
            df = df.append(pd.Series(info), ignore_index=True)
            df.to_excel('图书.xlsx', encoding='utf8', index=False)
            self.messagebox = QMessageBox()
            self.messagebox.information(self.messagebox, '欢迎使用', '保存成功')
        except Exception as e:
            with open('日志.csv',mode='a')as f:
                f.write(str(e))
                f.write('\n')

    def out_windows(self):
        self.out_window.show()

    def out(self):
        self.main_window.close()
        self.out_window.close()

    def close_out_window(self):
        self.out_window.close()

    def statistics(self):
        categary = self.main_window.comboBox.currentText()
        df = pd.read_excel('图书.xlsx').groupby('类别')
        for s, j in df:
            if s == categary:
                self.main_window.textBrowser_3.append(str(len(j['类别'])))
                df = j
        print(df)
        # print(categary)
        for i in range(1, len(df['类别']) + 1):
            id = (list(df['id'])[i - 1])

            newitem_i = QTableWidgetItem(str(id))
            self.statistic_window.tableWidget_2.setItem(i, 0, newitem_i)
        for i in range(1, len(df['类别']) + 1):
            z = (list(df['书名'])[i - 1])

            newitem_z = QTableWidgetItem(str(z))
            self.statistic_window.tableWidget_2.setItem(i, 1, newitem_z)
        for i in range(1, len(df['类别']) + 1):
            auth = list(df['作者'])[i - 1]

            newitem_i = QTableWidgetItem(str(auth))
            self.statistic_window.tableWidget_2.setItem(i, 2, newitem_i)
        for i in range(1, len(df['类别']) + 1):
            price = (list(df['价格'])[i - 1])

            newitem_i = QTableWidgetItem(str(price))
            self.statistic_window.tableWidget_2.setItem(i, 3, newitem_i)
        for i in range(1, len(df['类别']) + 1):
            fro = (list(df['出版单位'])[i - 1])

            newitem_i = QTableWidgetItem(str(fro))
            self.statistic_window.tableWidget_2.setItem(i, 4, newitem_i)
        for i in range(1, len(df['类别']) + 1):
            time = (list(df['类别'])[i - 1])

            newitem_i = QTableWidgetItem(str(time))
            self.statistic_window.tableWidget_2.setItem(i, 5, newitem_i)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    windows = Baidui_tr()
    ui = windows.main_window
    ui.show()
    sys.exit(app.exec_())
