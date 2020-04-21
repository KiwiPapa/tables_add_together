#conding=utf-8

import sys

from PyQt5.QtCore import QRect
from PyQt5.QtWidgets import QWidget, QApplication, QPushButton, QColorDialog, QFontDialog, \
    QTextEdit, QFileDialog, QDialog, QLineEdit, QLabel
from PyQt5.QtPrintSupport import QPageSetupDialog, QPrintDialog, QPrinter
from PyQt5 import QtCore, QtGui, QtWidgets
import numpy as np
import os, openpyxl
import pandas as pd
from changeOffice import Change

def get_thickness(x):
    thickness = x['井段End'] - x['井段Start']
    return thickness


def mkdir(path):
    path = path.strip()  # 去除首位空格
    path = path.rstrip("\\")  # 去除尾部 \ 符号
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        print(path + ' 创建成功')
        return True
    else:
        print(path + ' 目录已存在')
        return False

class AddTables(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setGeometry(300, 300, 500, 280)
        self.setWindowTitle('单层评价表合并并统计小程序')

        self.label1 = QLabel(self)
        self.label1.setGeometry(250, 5, 60, 20)
        self.label1.setText("拼接深度")
        self.lineEdit1 = QLineEdit(self)
        self.lineEdit1.setGeometry(250, 30, 60, 20)

        self.label2 = QLabel(self)
        self.label2.setGeometry(330, 5, 60, 20)
        self.label2.setText("开始深度")
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setGeometry(330, 30, 60, 20)

        self.label3 = QLabel(self)
        self.label3.setGeometry(410, 5, 60, 20)
        self.label3.setText("结束深度")
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setGeometry(410, 30, 60, 20)

        self.tx1 = QTextEdit(self)
        self.tx1.setGeometry(20, 85, 450, 60)

        self.tx2 = QTextEdit(self)
        self.tx2.setGeometry(20, 175, 450, 60)

        self.bt1 = QPushButton('选择上段表格', self)
        self.bt1.move(360, 60)
        self.bt3 = QPushButton('选择下段表格', self)
        self.bt3.move(360, 150)
        self.bt2 = QPushButton('填写完毕->运行', self)
        self.bt2.setGeometry(0, 0, 160, 60)
        self.bt2.move(20, 10)

        self.bt1.clicked.connect(self.openfiles1)
        self.bt3.clicked.connect(self.openfiles2)
        self.bt2.clicked.connect(self.run)
        
        self.show()

    def openfiles1(self):
        fnames = QFileDialog.getOpenFileNames(self, '打开第一个文件', './')#注意这里返回值是元组
        if fnames[0]:
            for fname in fnames[0]:
                self.tx1.append(fname)

    def openfiles2(self):
        fnames = QFileDialog.getOpenFileNames(self, '打开第二个文件', './')#注意这里返回值是元组
        if fnames[0]:
            for fname in fnames[0]:
                self.tx2.append(fname)
            
    def run(self):
        splicing_Depth = float(self.lineEdit1.text())

        # PATH = ".\\单层评价表"
        # for fileName in os.listdir(PATH):
        #     if '1单-1' in fileName:
        #         fileDir1 = PATH + "\\" + fileName
        #     elif '1单-2' in fileName:
        #         fileDir2 = PATH + "\\" + fileName

        fileDir1 = self.tx1.toPlainText()
        fileDir2 = self.tx2.toPlainText()

        df1 = pd.read_excel(fileDir1, header=2, index='序号')
        df1.drop([0], inplace=True)
        df1.loc[:, '井 段\n (m)'] = df1['井 段\n (m)'].str.replace(' ', '')  # 消除数据中空格
        df1.drop([len(df1)], inplace=True)
        df1['井段Start'] = df1['井 段\n (m)'].map(lambda x: x.split("-")[0])
        df1['井段End'] = df1['井 段\n (m)'].map(lambda x: x.split("-")[1])
        # 表格数据清洗
        df1.loc[:, "井段Start"] = df1["井段Start"].str.replace(" ", "").astype('float')
        df1.loc[:, "井段End"] = df1["井段End"].str.replace(" ", "").astype('float')

        # 截取拼接点以上的数据体
        df_temp1 = df1.loc[(df1['井段Start'] <= splicing_Depth), :].copy()  # 加上copy()可防止直接修改df报错
        # print(df_temp1)

        #####################################################
        df2 = pd.read_excel(fileDir2, header=2, index='序号')
        df2.drop([0], inplace=True)
        df2.loc[:, '井 段\n (m)'] = df2['井 段\n (m)'].str.replace(' ', '')  # 消除数据中空格
        df2.drop([len(df2)], inplace=True)
        df2['井段Start'] = df2['井 段\n (m)'].map(lambda x: x.split("-")[0])
        df2['井段End'] = df2['井 段\n (m)'].map(lambda x: x.split("-")[1])
        # 表格数据清洗
        df2.loc[:, "井段Start"] = df2["井段Start"].str.replace(" ", "").astype('float')
        df2.loc[:, "井段End"] = df2["井段End"].str.replace(" ", "").astype('float')

        # 截取拼接点以下的数据体
        df_temp2 = df2.loc[(df2['井段End'] >= splicing_Depth), :].copy()  # 加上copy()可防止直接修改df报错
        df_temp2.reset_index(drop=True, inplace=True)  # 重新设置列索引

        # print(df_temp2)

        df_all = df_temp1.append(df_temp2)
        df_all.reset_index(drop=True, inplace=True)  # 重新设置列索引
        # 对df_all进行操作
        df_all.loc[len(df_temp1) - 1, '井 段\n (m)'] = ''.join([str(df_all.loc[len(df_temp1) - 1, '井段Start']), '-', \
                                                              str(df_all.loc[len(df_temp1), '井段End'])])
        df_all.loc[len(df_temp1) - 1, '厚 度\n (m)'] = df_all.loc[len(df_temp1), '井段End'] - df_all.loc[
            len(df_temp1) - 1, '井段Start']
        df_all.loc[len(df_temp1) - 1, '最大声幅\n （%）'] = max(df_all.loc[len(df_temp1), '最大声幅\n （%）'], \
                                                             df_all.loc[len(df_temp1) - 1, '最大声幅\n （%）'])
        df_all.loc[len(df_temp1) - 1, '最小声幅\n  (%)'] = min(df_all.loc[len(df_temp1), '最小声幅\n  (%)'], \
                                                              df_all.loc[len(df_temp1) - 1, '最小声幅\n  (%)'])
        df_all.loc[len(df_temp1) - 1, '平均声幅\n  （%）'] = np.add(df_all.loc[len(df_temp1), '平均声幅\n  （%）'], \
                                                              df_all.loc[len(df_temp1) - 1, '平均声幅\n  （%）']) / 2

        df_all.drop(len(df_temp1), inplace=True)
        df_all.set_index(["解释\n序号"], inplace=True)
        df_all.reset_index(drop=True, inplace=True)  # 重新设置列索引
        # print(df_all)

        #################################################################
        # 在指定深度段统计

        # calculation_Start = float(input('请输入开始统计深度'))
        # calculation_End = float(input('请输入结束统计深度'))
        calculation_Start = float(self.lineEdit2.text())
        calculation_End = float(self.lineEdit3.text())

        start_Evaluation = df_all.loc[0, '井 段\n (m)'].split('-')[0]
        end_Evaluation = df_all.loc[len(df_all) - 1, '井 段\n (m)'].split('-')[1]
        if (calculation_End <= float(end_Evaluation)) & (calculation_Start >= float(start_Evaluation)):
            df_temp = df_all.loc[(df_all['井段Start'] >= calculation_Start) & (df_all['井段Start'] <= calculation_End), :]
            # 获取起始深度到第一层井段底界的结论
            df_temp1 = df_all.loc[(df_all['井段Start'] <= calculation_Start), :]
            start_to_upper_result = df_temp1.loc[len(df_temp1) - 1, '结论']
            # 补充储层界到井段的深度
            x, y = df_temp.shape
            df_temp = df_temp.reset_index()
            df_temp.drop(['index'], axis=1, inplace=True)
            if x >= 1:
                first_layer_start = df_temp.loc[0, '井段Start']
                upper = pd.DataFrame({'序号': '空',
                                      '井段': '空',
                                      '厚度': '空',
                                      '结论': start_to_upper_result,
                                      '井段Start': calculation_Start,
                                      '井段End': first_layer_start},
                                     index=[1])  # 自定义索引为：1 ，这里也可以不设置index
                df_temp.loc[len(df_temp) - 1, '井段End'] = calculation_End
                df_temp = pd.concat([upper, df_temp], ignore_index=True)
                # df_temp = df_temp.append(new, ignore_index=True)  # ignore_index=True,表示不按原来的索引，从0开始自动递增
                # print(df_temp)
            else:  # 储层包含在一个井段内的情况
                df_temp = pd.DataFrame({'序号': '空',
                                        '井段': '空',
                                        '厚度': '空',
                                        '结论': start_to_upper_result,
                                        '井段Start': calculation_Start,
                                        '井段End': calculation_End},
                                       index=[1])  # 自定义索引为：1 ，这里也可以不设置index
            df_temp.loc[:, "重计算厚度"] = df_temp.apply(get_thickness, axis=1)
            # print(df_temp)
            ratio_Series = df_temp.groupby(by=['结论'])['重计算厚度'].sum() / df_temp['重计算厚度'].sum() * 100
            if ratio_Series.__len__() == 2:
                if '好' not in ratio_Series:
                    ratio_Series = ratio_Series.append(pd.Series({'好': 0}))
                elif '中' not in ratio_Series:
                    ratio_Series = ratio_Series.append(pd.Series({'中': 0}))
                elif '差' not in ratio_Series:
                    ratio_Series = ratio_Series.append(pd.Series({'差': 0}))
            elif ratio_Series.__len__() == 1:
                if ('好' not in ratio_Series) & ('中' not in ratio_Series):
                    ratio_Series = ratio_Series.append(pd.Series({'好': 0}))
                    ratio_Series = ratio_Series.append(pd.Series({'中': 0}))
                elif ('好' not in ratio_Series) & ('差' not in ratio_Series):
                    ratio_Series = ratio_Series.append(pd.Series({'好': 0}))
                    ratio_Series = ratio_Series.append(pd.Series({'差': 0}))
                elif ('中' not in ratio_Series) & ('差' not in ratio_Series):
                    ratio_Series = ratio_Series.append(pd.Series({'中': 0}))
                    ratio_Series = ratio_Series.append(pd.Series({'差': 0}))

        # 统计结论
        actual_Hao = str(round((calculation_End - calculation_Start) * (ratio_Series['好'] / 100), 2))
        Hao_Ratio = str(round(ratio_Series['好'], 2))

        actual_Zhong = str(round((calculation_End - calculation_Start) * (ratio_Series['中'] / 100), 2))
        Zhong_Ratio = str(round(ratio_Series['中'], 2))

        actual_Cha = str(round(calculation_End - calculation_Start - float(actual_Hao) - float(actual_Zhong), 2))
        Cha_Ratio = str(round(ratio_Series['差'], 2))

        PATH = '.\\resources\\'
        wb = openpyxl.load_workbook(PATH + '1统模板.xlsx')
        sheet = wb[wb.sheetnames[0]]
        sheet['A1'] = ''.join(['第一界面水泥胶结统计表（', str(calculation_Start), '-', str(calculation_End), 'm）'])
        sheet['C4'] = actual_Hao
        sheet['D4'] = Hao_Ratio
        sheet['C5'] = actual_Zhong
        sheet['D5'] = Zhong_Ratio
        sheet['C6'] = actual_Cha
        sheet['D6'] = Cha_Ratio

        mkdir('.\\#DataOut')
        wb.save('.\\#DataOut\\解释成果表-1统.xlsx')

        # 单层统计表保存为Excel
        df_all.drop(['井段Start', '井段End'], axis=1, inplace=True)
        df_all.index = df_all.index + 1
        writer = pd.ExcelWriter('.\\#DataOut\\单层评价表—合并.xlsx')
        df_all.to_excel(writer, 'Sheet1')
        writer.save()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ad = AddTables()
    sys.exit(app.exec_())