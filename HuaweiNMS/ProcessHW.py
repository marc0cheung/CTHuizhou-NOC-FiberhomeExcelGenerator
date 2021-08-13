# -*- coding:utf-8 -*-
__author__ = "Marco Cheung"

import re
import sys
import time

import pandas
import xlrd
import xlwt
from PySide2 import QtCore, QtWidgets
from xlutils.copy import copy
from pyautogui import alert


class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.Step1Btn = QtWidgets.QPushButton("Generate Huawei OMS Sheet")
        self.Step2Btn = QtWidgets.QPushButton("Generate Huawei OPS Sheet")
        self.genButton = QtWidgets.QPushButton("Generate!")
        self.exitButton = QtWidgets.QPushButton("Exit")
        self.text = QtWidgets.QLabel("自动生成华为网管运维表格 V1.0\n\nby 张梓扬 Marco Cheung\n\n2021年8月 第一版\n\n请自行取用软件根目录下的文件")
        self.text.setAlignment(QtCore.Qt.AlignCenter)

        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.text)
        self.layout.addWidget(self.Step1Btn)
        self.layout.addWidget(self.Step2Btn)
        self.layout.addWidget(self.genButton)
        self.layout.addWidget(self.exitButton)
        self.setLayout(self.layout)

        self.Step1Btn.clicked.connect(genOMS)
        self.Step2Btn.clicked.connect(genOPS)
        self.genButton.clicked.connect(wholeProcess)
        self.exitButton.clicked.connect(app.exit)


def genOMS():
    time_start = time.time()

    hw_Ori = pandas.read_excel("当前性能数据.xls", sheet_name=0, header=7)
    newOMS = xlwt.Workbook()

    # 化简原始数据表格, 减少遍历次数
    hw_OriSimple = hw_Ori.copy(deep=True)
    for i in range(0, len(hw_Ori)):
        if '激光器' in hw_Ori.iat[i, 1]:
            hw_OriSimple.drop([i, i], inplace=True)
    # 重置 DataFrame 'hw_Ori' 的索引
    hw_OriSimple.reset_index(drop=True, inplace=True)
    print("genOMS: 已完成hw_Ori简化，去除含“激光器”的内容")

    sheet = newOMS.add_sheet('华为波分OMS检查（每月）', cell_overwrite_ok=True)
    sheet_title = ['网元名', '方向', '方向/槽位', '输入光功率', '输出光功率', 'VOA', '建议处理', '备注']

    # 写入表格标题
    for i in range(0, len(sheet_title)):
        sheet.write(0, i, sheet_title[i])

    # 写入 方向/槽位 信息
    for i in range(0, len(hw_OriSimple)):
        sheet.write(i + 1, 2, hw_OriSimple.iat[i, 0])
    print("genOMS: 完成写入 表格标题 和 方向/槽位 信息")

    # 写入 输入光功率 和 输出光功率
    i = 0
    while i <= len(hw_OriSimple)-1:
        sheet.write_merge(i + 1, i + 2, 3, 3, hw_OriSimple.iat[i, 5])  # 写入 输入光功率
        i = i + 2
    print('genOMS: 完成写入 输入光功率')

    # 写入 输出光功率
    i = 1
    while i <= len(hw_OriSimple)-1:
        sheet.write_merge(i, i + 1, 4, 4, hw_OriSimple.iat[i, 5])  # 写入 输出光功率
        i = i + 2
    print('genOMS: 完成写入 输出光功率')

    # 合并 VOA 单元格，方便后期人工核查 VOA 数据
    i = 0
    while i <= len(hw_OriSimple)-1:
        sheet.write_merge(i + 1, i + 2, 5, 5, '')  # Merge "VOA"
        i = i + 2

    newOMS.save('genHWOMS.xls')
    time_end = time.time()
    print("genOMS: DONE, Total Time Cost: ", time_end - time_start)
    alert(text="Targeted Excel(OMS) Generated!\n点击“好的”将继续生成OPS表\n用时："+str(time_end - time_start), title="处理结果", button="好的")


def genOPS():
    time_start = time.time()

    hw_Ori = pandas.read_excel("当前性能数据.xls", sheet_name=0, header=7)
    HWops_template = pandas.read_excel("HWops_template.xlsx", header=0, sheet_name=0)

    # 把合并的单元格内容分配到每一行
    HWops_template['波分环'] = HWops_template['波分环'].ffill()
    HWops_template['双纤衰耗差值（dBm）'] = HWops_template['双纤衰耗差值（dBm）'].ffill()
    # ops_template['光路编码'] = ops_template['光路编码'].ffill()
    HWops_template['是否检修（一级）【K列判决】'] = HWops_template['是否检修（一级）【K列判决】'].ffill()
    HWops_template['是否检修（二级）【K&J列判决】'] = HWops_template['是否检修（二级）【K&J列判决】'].ffill()
    HWops_template['理论光路长度（KM）'] = HWops_template['理论光路长度（KM）'].ffill()
    HWops_template['理论衰耗（dBm）'] = HWops_template['理论衰耗（dBm）'].ffill()

    # 化简原始数据表格, 减少遍历次数
    hw_OriSimple = hw_Ori.copy(deep=True)
    for i in range(0, len(hw_Ori)):
        if '激光器' not in hw_Ori.iat[i, 1]:
            hw_OriSimple.drop([i, i], inplace=True)
    # 重置 DataFrame 'hw_Ori' 的索引
    hw_OriSimple.reset_index(drop=True, inplace=True)
    print("genOSC: 已完成hw_Ori简化，去除不含“激光器”的内容")

    # 写入输出光功率
    for i in range(0, len(HWops_template)):
        for j in range(0, len(hw_OriSimple)):
            if (HWops_template.iat[i, 1] in hw_OriSimple.iat[j, 0]) and re.split('-', HWops_template.iat[i, 4])[-1] in hw_OriSimple.iat[j, 0]:
                if '输出' in hw_OriSimple.iat[j, 1]:
                    HWops_template.iat[i, 2] = hw_OriSimple.iat[j, 5]
                else:
                    continue
            else:
                continue
    print("genOSC: Pandas 完成写入 “输出光功率” ")

    # 写入输入光功率
    for i in range(0, len(HWops_template)):
        for j in range(0, len(hw_OriSimple)):
            if (HWops_template.iat[i, 4] in hw_OriSimple.iat[j, 0]) and re.split('-', HWops_template.iat[i, 1])[-1] in hw_OriSimple.iat[j, 0]:
                if '输入' in hw_OriSimple.iat[j, 1]:
                    HWops_template.iat[i, 3] = hw_OriSimple.iat[j, 5]
                else:
                    continue
            else:
                continue
    print("genOSC: Pandas 完成写入 “输入光功率” ")

    # 计算并写入A-B衰耗、实际与理论衰耗差值
    for i in range(0, len(HWops_template)):
        HWops_template.iat[i, 5] = abs(HWops_template.iat[i, 2] - HWops_template.iat[i, 3])
        HWops_template.iat[i, 9] = abs(HWops_template.iat[i, 5] - HWops_template.iat[i, 8])

    # 写入 “双纤衰耗差值” 和 “是否检修（一级）【K列判决】”
    i = 0
    while i <= len(HWops_template)-1:
        HWops_template.iat[i, 10] = abs(HWops_template.iat[i, 5] - HWops_template.iat[i + 1, 5])
        HWops_template.iat[i+1, 10] = abs(HWops_template.iat[i, 5] - HWops_template.iat[i + 1, 5])
        if abs(HWops_template.iat[i, 5] - HWops_template.iat[i + 1, 5]) > 5:
            HWops_template.iat[i, 11] = '是'
            HWops_template.iat[i+1, 11] = '是'
        else:
            HWops_template.iat[i, 11] = '否'
            HWops_template.iat[i+1, 11] = '否'
        i = i + 2
    print("genOSC: Pandas 完成写入”双纤衰耗差值“和”一级检修判决“")

    # 写入 “是否检修（二级）【K&J列判决】”
    i = 0
    while i <= len(HWops_template) - 1:
        if (abs(HWops_template.iat[i, 5] - HWops_template.iat[i + 1, 5]) > 5) or (abs(HWops_template.iat[i, 5] - HWops_template.iat[i, 8]) > 5) or (abs(HWops_template.iat[i + 1, 5] - HWops_template.iat[i + 1, 8]) > 5):
            HWops_template.iat[i, 12] = '是'
            HWops_template.iat[i + 1, 12] = '是'
        else:
            HWops_template.iat[i, 12] = '否'
            HWops_template.iat[i + 1, 12] = '否'
        i = i + 2
    print('genNewOSC: Pandas 完成写入 “是否检修（二级）【K&J列判决】”')

    # 利用Pandas实现分Sheet写入表格
    with pandas.ExcelWriter('./genHWOPS.xls') as writer:
        HWops_template.to_excel(writer, encoding='utf-8', sheet_name='华为波分OPS检查（每月）', index=False)
    print("genOSC: Pandas 已保存初始文档，交由 xlwt 处理")

    r_xls = xlrd.open_workbook("genHWOPS.xls")  # 读取excel文件
    excelCopy = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
    sheet3 = excelCopy.get_sheet(0)
    print("genOSC: XLRD 已读取，产生XLWT输入")

    # 合并原表中需要合并的单元格
    i = 0
    while i <= len(HWops_template) - 1:
        sheet3.write_merge(i + 1, i + 2, 7, 7, HWops_template.iat[i, 7])
        sheet3.write_merge(i + 1, i + 2, 8, 8, HWops_template.iat[i, 8])
        sheet3.write_merge(i + 1, i + 2, 10, 10, HWops_template.iat[i, 10])
        sheet3.write_merge(i + 1, i + 2, 11, 11, HWops_template.iat[i, 11])
        sheet3.write_merge(i + 1, i + 2, 12, 12, HWops_template.iat[i, 12])
        i = i + 2

    excelCopy.save('genHWOPS.xls')
    time_end = time.time()
    print("gen1: DONE, Total Time Cost: ", time_end - time_start)
    alert(text="Targeted Excel Generated!\n用时："+str(time_end - time_start), title="处理结果", button="好的")


def wholeProcess():
    genOMS()
    genOPS()


if __name__ == "__main__":
    app = QtWidgets.QApplication([])

    widget = MyWidget()
    widget.resize(500, 600)
    widget.setWindowTitle("华为网管表格自动生成工具 V1.0")
    widget.show()

    sys.exit(app.exec_())