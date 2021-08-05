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

        self.Step1Btn = QtWidgets.QPushButton("Generate OMS Sheet using oa.xlsx")
        self.Step2Btn = QtWidgets.QPushButton("Generate OCH Sheet using ocp.xlsx")
        self.Step3Btn = QtWidgets.QPushButton("Generate OPS Sheet using osc.xlsx")
        self.genButton = QtWidgets.QPushButton("Generate!")
        self.exitButton = QtWidgets.QPushButton("Exit")
        self.text = QtWidgets.QLabel("自动生成烽火系统运维表格 V1.0\n\nby 张梓扬 Marco Cheung\n\n2021年8月 第一版\n\n请自行合并软件根目录下的三个文件")
        self.text.setAlignment(QtCore.Qt.AlignCenter)

        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.text)
        self.layout.addWidget(self.Step1Btn)
        self.layout.addWidget(self.Step2Btn)
        self.layout.addWidget(self.Step3Btn)
        self.layout.addWidget(self.genButton)
        self.layout.addWidget(self.exitButton)
        self.setLayout(self.layout)

        self.Step1Btn.clicked.connect(genNewOA)
        self.Step2Btn.clicked.connect(genNewOCP)
        self.Step3Btn.clicked.connect(genNewNewOSC)
        self.genButton.clicked.connect(wholeProcess)
        self.exitButton.clicked.connect(app.exit)


def genNewOA():
    time_start = time.time()
    oa_Ori = pandas.read_excel("oa.xlsx", sheet_name=0, header=0)
    newOA = xlwt.Workbook()
    # ====================================== 执行正则匹配并化简多余项 ====================================================
    oaDirectSlotMatch = []
    oa_nameReg = "\:"
    oa_noReg_1 = '\['
    oa_noReg_2 = "\:\:"

    for i in range(0, len(oa_Ori)):
        oaDirectSlotMatch.append(
            re.split(oa_nameReg, oa_Ori[1][i])[-1] + ":" + re.split(oa_noReg_1, oa_Ori[2][i])[0] + ":" +
            re.split(oa_noReg_2, oa_Ori[2][i])[1])
    oaDirectSlotMatch = sorted(set(oaDirectSlotMatch), key=oaDirectSlotMatch.index)

    sheet1 = newOA.add_sheet('烽火波分OMS检查（每月）', cell_overwrite_ok=True)
    sheet1_title = ["网元名", "方向/槽位", "输入光功率", "输出光功率", "VOA", "处理建议", "备注"]

    # =================================== 写入表格 ==========================================
    # 写入表格标题
    for i in range(0, len(sheet1_title)):
        sheet1.write(0, i, sheet1_title[i])

    # 写入OA的方向/槽位信息
    for i in range(0, len(oaDirectSlotMatch)):
        sheet1.write(i + 1, 1, oaDirectSlotMatch[i])

    # 通过检索实现数据的对应输入
    for i in range(0, len(oa_Ori)):
        index = oaDirectSlotMatch.index(
            re.split(oa_nameReg, oa_Ori[1][i])[-1] + ":" + re.split(oa_noReg_1, oa_Ori[2][i])[0] + ":" +
            re.split(oa_noReg_2, oa_Ori[2][i])[1])
        if oa_Ori[6][i] == "IOP":
            sheet1.write(index + 1, 2, oa_Ori[7][i])
        elif oa_Ori[6][i] == "OOP":
            sheet1.write(index + 1, 3, oa_Ori[7][i])
        elif oa_Ori[6][i] == "VOA_ATT":
            sheet1.write(index + 1, 4, oa_Ori[7][i])
        else:
            continue

    newOA.save("genNewOA.xls")
    newOA = pandas.read_excel("genNewOA.xls", sheet_name="烽火波分OMS检查（每月）", header=0)
    newOA.sort_values(by=['方向/槽位'], inplace=True)
    with pandas.ExcelWriter('./genNewOA.xls') as writer:
        newOA.to_excel(writer, encoding='utf-8', sheet_name='烽火波分OMS检查（每月）', index=False)

    time_end = time.time()
    print("genNewOA: DONE, Total Time Cost: ", time_end - time_start)
    alert(text="Targeted Excel Generated!\n用时：" + str(time_end - time_start), title="处理结果", button="好的")


def genNewOCP():
    time_start = time.time()
    ocp_Ori = pandas.read_excel("ocp.xlsx", sheet_name=0, header=0)
    newOCP = xlwt.Workbook()
    # ====================================== 执行正则匹配并化简多余项 ====================================================
    ocpDirectSlotMatch = []
    ocp_nameReg = "\:"
    ocp_noReg_1 = '\['
    ocp_noReg_2 = "\:\:"

    for i in range(0, len(ocp_Ori)):
        ocpDirectSlotMatch.append(
            re.split(ocp_nameReg, ocp_Ori[1][i])[-1] + ":" + re.split(ocp_noReg_1, ocp_Ori[2][i])[0] + ":" +
            re.split(ocp_noReg_2, ocp_Ori[2][i])[1])
    ocpDirectSlotMatch = sorted(set(ocpDirectSlotMatch), key=ocpDirectSlotMatch.index)

    sheet2 = newOCP.add_sheet('烽火波分OCH光功率检查（每月）', cell_overwrite_ok=True)
    sheet2_title = ["烽火波分环", "方向/板卡/端口", "输入光功率", "处理建议", "备注", "差异"]

    # =================================== 写入表格 ==========================================
    # 写入表格标题
    for i in range(0, len(sheet2_title)):
        sheet2.write(0, i, sheet2_title[i])

    # 写入OCP的方向/板卡/端口信息
    for i in range(0, len(ocpDirectSlotMatch)):
        sheet2.write(i + 1, 1, ocpDirectSlotMatch[i])

    # 通过检索实现数据的对应输入
    reg = "[-+]?\d+.?\d*"  # Rule for Matching Numbers
    for i in range(0, len(ocp_Ori)):
        index = ocpDirectSlotMatch.index(
            re.split(ocp_nameReg, ocp_Ori[1][i])[-1] + ":" + re.split(ocp_noReg_1, ocp_Ori[2][i])[0] + ":" +
            re.split(ocp_noReg_2, ocp_Ori[2][i])[1])
        if ocp_Ori[6][i] == "IOP":
            if ocp_Ori[7][i] == '收无光':
                sheet2.write(index + 1, 2, "收无光")
            else:
                sheet2.write(index + 1, 2, float(re.findall(reg, str(ocp_Ori[7][i]))[0]))
        else:
            continue

    newOCP.save("genNewOCP.xls")

    # 读入表格进行排序
    newOCP = pandas.read_excel("genNewOCP.xls", sheet_name="烽火波分OCH光功率检查（每月）", header=0)
    newOCP.sort_values(by=['方向/板卡/端口'], inplace=True)
    newOCP.reset_index(drop=True, inplace=True)

    # 创建一个化简的ocp_OriSimple，用空间换时间
    ocp_OriSimple = ocp_Ori
    for i in range(0, len(ocp_OriSimple)):
        if (ocp_OriSimple[6][i] == 'IOP') or (ocp_OriSimple[6][i] == 'IOP_MAX'):
            ocp_OriSimple.drop([i, i], inplace=True)

    # 重置 DataFrame 'osc_Ori' 的索引
    ocp_OriSimple.reset_index(drop=True, inplace=True)
    ocp_Ori = pandas.read_excel("ocp.xlsx", sheet_name=0, header=0)
    print("genNewOCP: 已完成ocp_Ori简化，去除含IOP、IOP_MAX内容")

    # 对于没有IOP的线路，选用IOP_MIN进行填充
    for i in range(0, len(newOCP)):
        if str(newOCP['输入光功率'][i]) == 'nan':
            for j in range(0, len(ocp_OriSimple)):
                dest = re.split(ocp_nameReg, ocp_OriSimple[1][j])[-1] + ":" + re.split(ocp_noReg_1, ocp_OriSimple[2][j])[0] + ":" + re.split(ocp_noReg_2, ocp_OriSimple[2][j])[1]
                if newOCP.iat[i, 1] == dest and ocp_OriSimple[6][j] == "IOP_MIN":
                    if ocp_OriSimple[7][j] == '收无光':
                        newOCP.iat[i, 2] = '收无光'
                    else:
                        newOCP.iat[i, 2] = float(re.findall(reg, str(ocp_OriSimple[7][j]))[0])
                else:
                    continue

    with pandas.ExcelWriter('./genNewOCP.xls') as writer:
        newOCP.to_excel(writer, encoding='utf-8', sheet_name='烽火波分OCH光功率检查（每月）', index=False)

    # 合并单元格计算差异
    r_xls = xlrd.open_workbook("genNewOCP.xls")  # 读取excel文件
    excelCopy = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
    sheet2 = excelCopy.get_sheet(0)
    OCPData = pandas.read_excel("genNewOCP.xls", sheet_name="烽火波分OCH光功率检查（每月）", header=0, usecols=[2, 2]).values
    OCPName = pandas.read_excel("genNewOCP.xls", sheet_name="烽火波分OCH光功率检查（每月）", header=0, usecols=[1, 1])

    i = 0
    while i <= len(newOCP) - 1:
        # OCPData_Diff.append(abs(OCPData[i] - OCPData[i+1]))
        # sheet2.write_merge(i+1, i+2, 5, 5, "")
        if str(OCPData[i]) == "['收无光']" or str(OCPData[i + 1]) == "['收无光']":
            sheet2.write_merge(i + 1, i + 2, 5, 5, "无法计算")
        else:
            sheet2.write_merge(i + 1, i + 2, 5, 5, float(abs(float(OCPData[i]) - float(OCPData[i + 1]))))
        i = i + 2

    # 把错误排序的节点标红
    repeat_reg = "[T][R][X][AB]"

    pattern = xlwt.Pattern()  # Create the pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = 5

    style = xlwt.XFStyle()  # Create the pattern
    style.pattern = pattern  # Add pattern to style

    i = 0
    while i <= len(ocpDirectSlotMatch) - 1:
        if re.findall(repeat_reg, OCPName["方向/板卡/端口"][i]) == re.findall(repeat_reg, OCPName["方向/板卡/端口"][i + 1]):
            sheet2.write(i + 1, 1, OCPName["方向/板卡/端口"][i], style=style)
            i = i + 2
        else:
            i = i + 2

    print('genNewOCP: 对错误节点的标注已完成')
    excelCopy.save("genNewOCP.xls")
    time_end = time.time()
    print("genNewOCP: DONE, Total Time Cost: ", time_end - time_start)
    alert(text="Targeted Excel Generated!\n用时：" + str(time_end - time_start), title="处理结果", button="好的")

'''
def genNewOSC():
    osc_Ori = pandas.read_excel("osc.xlsx", header=0, sheet_name=0)
    newOSC = xlwt.Workbook()
    sheet3 = newOSC.add_sheet('烽火波分OPS检查（每月）', cell_overwrite_ok=True)
    sheet3_title = ["波分环", "A节点设备名称", "OCS-输出光功率（dBm）", "OCS-输入光功率（dBm）", "B节点设备名称", "A->B衰耗"]
    oscSum = []
    oscName = []
    for i in range(0, len(osc_Ori)):
        if str(osc_Ori[3][i]) == 'nan':
            oscSum.append(osc_Ori[1][i] + ":" + osc_Ori[2][i] + ":" + "")
        else:
            oscSum.append(osc_Ori[1][i] + ":" + osc_Ori[2][i] + ":" + osc_Ori[3][i])
    oscSum = sorted(set(oscSum), key=oscSum.index)

    reg = "(\d{2}-\d*(-|)[\u4e00-\u9fa5]*(OA|OTM-[\u4e00-\u9fa5]*|OTM|-ROADM\(OA\)|ROADM\(OA\)|-ROADM))"
    regOTMFX = "\d{2}\-\d*-[\u4e00-\u9fa5]*OTM-[\u4e00-\u9fa5]*方向|\d{2}\-\d*[\u4e00-\u9fa5]*OTM-[\u4e00-\u9fa5]*方向"
    OTMFX = []

    for i in range(0, len(oscSum)):
        if re.findall(regOTMFX, oscSum[i]) == []:
            oscName.append(re.findall(reg, oscSum[i])[0][0])
        else:
            OTMFX.append(re.findall(reg, oscSum[i])[0][0])
    # OTM-XX方向都是只重复一次的，所以可以先提取出来然后去掉重复元素
    OTMFX = sorted(set(OTMFX), key=OTMFX.index)

    # =================================== 写入表格 ==========================================
    # 写入表格标题
    for i in range(0, len(sheet3_title)):
        sheet3.write(0, i, sheet3_title[i])

    # 写入OTM-XX方向这一类的线路名称
    for i in range(0, len(OTMFX)):
        sheet3.write(i + 1, 1, OTMFX[i])

    # 检索和A节点对应的B节点
    regNum = "[-+]?\d+.?\d*"  # Rule for Matching Numbers
    for i in range(0, len(OTMFX)):
        knot_A = re.split("OTM-", re.split("惠州", OTMFX[i])[-1])[0]
        knot_B = re.split("方向", re.split("OTM-", re.split("惠州", OTMFX[i])[-1])[-1])[0]
        Num1 = re.split("-", OTMFX[79])[0]
        searchKey = Num1 + "惠州" + knot_B + "OTM-" + knot_A + "方向"
        searchResult = difflib.get_close_matches(searchKey, OTMFX, 1, cutoff=0.5)
        sheet3.write(i + 1, 4, searchResult)
        for j in range(0, len(osc_Ori)):
            if re.findall(reg, osc_Ori[1][j])[0][0] == searchResult[0] and re.split("\:\:", osc_Ori[2][j])[-1] == 'OSC_W' and osc_Ori[6][j] == 'IOP_MIN':
                # sheet3.write(i + 1, 3, osc_Ori[7][j])
                if osc_Ori[7][j] == '无光':
                    sheet3.write(i + 1, 3, osc_Ori[7][j])
                else:
                    sheet3.write(i + 1, 3, float(re.findall(regNum, osc_Ori[7][j])[0]))

    for j in range(0, len(OTMFX)):
        for i in range(0, len(osc_Ori)):
            if re.findall(reg, osc_Ori[1][i])[0][0] == OTMFX[j] and re.split("\:\:", osc_Ori[2][i])[-1] == 'OSC_W' and osc_Ori[6][i] == 'OOP_MIN':
                if osc_Ori[7][i] == '无光':
                    sheet3.write(i + 1, 3, osc_Ori[7][i])
                else:
                    sheet3.write(j + 1, 2, float(re.findall(regNum, osc_Ori[7][i])[0]))

    newOSC.save("genNewOSC.xls")
    newOSC = pandas.read_excel("genNewOSC.xls", sheet_name="烽火波分OPS检查（每月）", header=0)
    newOSC.sort_values(by=['A节点设备名称'], inplace=True)
    with pandas.ExcelWriter('./genNewOSC.xls') as writer:
        newOSC.to_excel(writer, encoding='utf-8', sheet_name='烽火波分OPS检查（每月）', index=False)

    r_xls = xlrd.open_workbook("genNewOSC.xls")  # 读取excel文件
    excelCopy = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
    sheet_minus = excelCopy.get_sheet(0)
    OSC_IOOP = pandas.read_excel("genNewOSC.xls", header=0, usecols=[2, 3])

    for i in range(0, len(OSC_IOOP)):
        if isinstance(OSC_IOOP["OCS-输出光功率（dBm）"][i], float) and isinstance(OSC_IOOP["OCS-输入光功率（dBm）"][i], float):
            sheet_minus.write(i + 1, 5, OSC_IOOP["OCS-输出光功率（dBm）"][i] - OSC_IOOP["OCS-输入光功率（dBm）"][i])
        else:
            sheet_minus.write(i + 1, 5, "无法计算")

    excelCopy.save("genNewOSC.xls")

    print("DONE")
'''


def genNewNewOSC():
    time_start = time.time()

    reg = "(\d{2}-\d*(-|)[\u4e00-\u9fa5]*(OA|OTM-[\u4e00-\u9fa5]*|OTM|-ROADM\(OA\)|ROADM\(OA\)|-ROADM))"
    regNum = "[-+]?\d+.?\d*"  # Rule for Matching Numbers

    ops_template = pandas.read_excel("ops_template.xlsx", header=3, sheet_name=0)
    osc_Ori = pandas.read_excel("osc.xlsx", header=0, sheet_name=0)

    # 删除无用的含GE的信息列，减少遍历次数
    for i in range(0, len(osc_Ori)):
        if 'GE' in osc_Ori[2][i]:
            osc_Ori.drop([i, i], inplace=True)
    # 重置 DataFrame 'osc_Ori' 的索引
    osc_Ori.reset_index(drop=True, inplace=True)
    print("genNewOSC: 已完成osc_Ori简化，去除含GE内容")

    # 把合并的单元格内容分配到每一行
    ops_template['环'] = ops_template['环'].ffill()
    ops_template['双纤衰耗差值（dBm）'] = ops_template['双纤衰耗差值（dBm）'].ffill()
    # ops_template['光路编码'] = ops_template['光路编码'].ffill()
    ops_template['是否检修（一级）【K列判决】'] = ops_template['是否检修（一级）【K列判决】'].ffill()
    ops_template['是否检修（二级）【K&J列判决】'] = ops_template['是否检修（二级）【K&J列判决】'].ffill()
    ops_template['理论光路长度（KM）'] = ops_template['理论光路长度（KM）'].ffill()
    ops_template['理论衰耗（dBm）'] = ops_template['理论衰耗（dBm）'].ffill()

    # 取得第一、第二、第三类的数目
    countOTMFX = 0
    countOTM = 0
    countROADM = 0
    for i in range(0, len(ops_template)):
        if re.fullmatch('(\d{2}-\d*(|-)[\u4e00-\u9fa5]*OTM-[\u4e00-\u9fa5]*方向)', ops_template.iat[i, 1]) is not None:
            countOTMFX = countOTMFX + 1
        elif re.fullmatch("(\d{2}-\d*-[\u4e00-\u9fa5]*(OTM|OA))", ops_template.iat[i, 1]) is not None:
            countOTM = countOTM + 1
        elif 'ROADM' in ops_template.iat[i, 1]:
            countROADM = countROADM + 1
        else:
            continue

    countOTM = countOTM + countOTMFX
    countROADM = countOTM + countROADM

    # 第一类：写入【NN-NN-惠州XXOTM-XX方向】这一类
    for i in range(0, countOTMFX):
        for j in range(0, len(osc_Ori)):
            if ops_template["A结点设备名称"][i] == re.findall(reg, osc_Ori[1][j])[0][0] and re.split("::", osc_Ori[2][j])[-1] == 'OSC_W' and osc_Ori[6][j] == 'OOP_MIN':
                ops_template.iat[i, 2] = float(re.findall(regNum, osc_Ori[7][j])[0])
                break
            else:
                continue

    for i in range(0, countOTMFX):
        for j in range(0, len(osc_Ori)):
            if ops_template["B结点设备名称"][i] == re.findall(reg, osc_Ori[1][j])[0][0] and re.split("::", osc_Ori[2][j])[-1] == 'OSC_W' and osc_Ori[6][j] == 'IOP_MIN':
                ops_template.iat[i, 3] = float(re.findall(regNum, osc_Ori[7][j])[0])
                ops_template.iat[i, 5] = ops_template.iat[i, 2] - ops_template.iat[i, 3]
                break
            else:
                continue
    print("genNewOSC: 第一类写入完成")

    # 第二类：写入【NN-NN-惠州XXOTM】这一类
    for i in range(countOTMFX, countOTM):
        for j in range(0, len(osc_Ori)):
            place_B = re.split("OTM|OA", re.split("惠州", ops_template['B结点设备名称'][i])[-1])[0]
            if ops_template["A结点设备名称"][i] == re.findall(reg, osc_Ori[1][j])[0][0] and osc_Ori[6][j] == 'OOP_MIN':
                if 'E' + place_B in osc_Ori[2][j]:
                    if re.split("::", osc_Ori[2][j])[-1] == 'OSC_E':
                        ops_template.iat[i, 2] = float(re.findall(regNum, osc_Ori[7][j])[0])
                        break
                elif 'W' + place_B in osc_Ori[2][j]:
                    if re.split("::", osc_Ori[2][j])[-1] == 'OSC_W':
                        ops_template.iat[i, 2] = float(re.findall(regNum, osc_Ori[7][j])[0])
                        break
                elif place_B in osc_Ori[3][j]:
                    ops_template.iat[i, 2] = float(re.findall(regNum, osc_Ori[7][j])[0])
                    break
                else:
                    continue
            else:
                continue

    for i in range(countOTMFX, countOTM):
        for j in range(0, len(osc_Ori)):
            place_A = re.split("OTM|OA", re.split("惠州", ops_template['A结点设备名称'][i])[-1])[0]
            if ops_template["B结点设备名称"][i] == re.findall(reg, osc_Ori[1][j])[0][0] and osc_Ori[6][j] == 'IOP_MIN':
                if 'E' + place_A in osc_Ori[2][j]:
                    if re.split("::", osc_Ori[2][j])[-1] == 'OSC_E':
                        ops_template.iat[i, 3] = float(re.findall(regNum, osc_Ori[7][j])[0])
                        ops_template.iat[i, 5] = ops_template.iat[i, 2] - ops_template.iat[i, 3]
                        break
                elif 'W' + place_A in osc_Ori[2][j]:
                    if re.split("::", osc_Ori[2][j])[-1] == 'OSC_W':
                        ops_template.iat[i, 3] = float(re.findall(regNum, osc_Ori[7][j])[0])
                        ops_template.iat[i, 5] = ops_template.iat[i, 2] - ops_template.iat[i, 3]
                        break
                elif place_A in osc_Ori[3][j]:
                    ops_template.iat[i, 3] = float(re.findall(regNum, osc_Ori[7][j])[0])
                    ops_template.iat[i, 5] = ops_template.iat[i, 2] - ops_template.iat[i, 3]
                    break
                else:
                    continue
            else:
                continue

    print("genNewOSC: 第二类写入完成")

    # 第三类：写入 NN-NN-惠州XXn-ROADM/-ROADM（OA）这一类
    osc_OriROADM = osc_Ori
    for i in range(0, len(osc_OriROADM)):
        if 'ROADM' not in osc_Ori[1][i]:
            osc_OriROADM.drop([i, i], inplace=True)

    osc_OriROADM.reset_index(drop=True, inplace=True)

    for i in range(countOTM, countROADM):
        for j in range(0, len(osc_OriROADM)):
            place_B = re.split("-ROADM(OA)|-ROADM|ROADM", re.split("惠州", ops_template['B结点设备名称'][i])[-1])[0] + '方向'
            if re.findall("\d{2}-\d*-[\u4e00-\u9fa5]*", ops_template["A结点设备名称"][i])[0] in osc_OriROADM[1][j] and osc_OriROADM[6][j] == 'OOP_MIN' and 'OSC_W' in osc_OriROADM[2][j]:
                if (place_B in osc_OriROADM[2][j]) or (place_B in str(osc_OriROADM[3][j])):
                    ops_template.iat[i, 2] = float(re.findall(regNum, osc_OriROADM[7][j])[0])
                    break
                else:
                    continue
            else:
                continue

    for i in range(countOTM, countROADM):
        for j in range(0, len(osc_OriROADM)):
            place_A = re.split("-ROADM(OA)|-ROADM|ROADM", re.split("惠州", ops_template['A结点设备名称'][i])[-1])[0] + '方向'
            if re.findall("\d{2}-\d*-[\u4e00-\u9fa5]*", ops_template["B结点设备名称"][i])[0] in osc_OriROADM[1][j] and osc_OriROADM[6][j] == 'IOP_MIN' and 'OSC_W' in osc_OriROADM[2][j]:
                if (place_A in osc_OriROADM[2][j]) or (place_A in str(osc_OriROADM[3][j])):
                    ops_template.iat[i, 3] = float(re.findall(regNum, osc_OriROADM[7][j])[0])
                    ops_template.iat[i, 5] = ops_template.iat[i, 2] - ops_template.iat[i, 3]
                    break
                else:
                    continue
            else:
                continue

    print("genNewOSC: 第三类写入完成")

    # 利用Pandas实现分Sheet写入表格
    with pandas.ExcelWriter('./genNewNewOSC.xls') as writer:
        ops_template.to_excel(writer, encoding='utf-8', sheet_name='烽火波分OPS检查（每月）', index=False)
    print("genNewOSC: Pandas 已保存初始文档，交由 xlwt 处理")

    r_xls = xlrd.open_workbook("genNewNewOSC.xls")  # 读取excel文件
    excelCopy = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
    sheet3 = excelCopy.get_sheet(0)
    print("genNewOSC: XLRD 已读取，产生XLWT输入")

    # 写入“实际与理论衰耗差值”
    for i in range(0, len(ops_template)):
        sheet3.write(i + 1, 9, abs(ops_template.iat[i, 5] - ops_template.iat[i, 8]))
    print("genNewOSC: XLWT完成写入”实际与理论衰耗差值“")

    # 写入 “双纤衰耗差值” 和 “是否检修（一级）【K列判决】”
    i = 0
    while i <= len(ops_template) - 1:
        sheet3.write_merge(i + 1, i + 2, 10, 10, abs(ops_template.iat[i, 5] - ops_template.iat[i + 1, 5]))
        if abs(ops_template.iat[i, 5] - ops_template.iat[i + 1, 5]) > 5:
            sheet3.write_merge(i + 1, i + 2, 11, 11, '是')
        else:
            sheet3.write_merge(i + 1, i + 2, 11, 11, '否')
        i = i + 2

    print("genNewOSC: XLWT完成写入 “双纤衰耗差值” 和 “是否检修（一级）【K列判决】”")

    # 写入 “是否检修（二级）【K&J列判决】”
    i = 0
    while i <= len(ops_template) - 1:
        if (abs(ops_template.iat[i, 5] - ops_template.iat[i + 1, 5]) > 5) or (abs(ops_template.iat[i, 5] - ops_template.iat[i, 8]) > 5) or (abs(ops_template.iat[i + 1, 5] - ops_template.iat[i + 1, 8]) > 5):
            sheet3.write_merge(i + 1, i + 2, 12, 12, '是')
        else:
            sheet3.write_merge(i + 1, i + 2, 12, 12, '否')
        i = i + 2

    print('genNewOSC: XLWT完成写入 “是否检修（二级）【K&J列判决】”')

    # 将Pandas拆分的光路编码、理论光路长度、理论衰耗进行合并
    i = 0
    while i <= len(ops_template) - 1:
        if str(ops_template.iat[i, 13]) != 'nan':
            sheet3.write_merge(i + 1, i + 2, 13, 13, ops_template.iat[i, 13])
        else:
            sheet3.write_merge(i + 1, i + 2, 13, 13, "")
        i = i + 2

    i = 0
    while i <= len(ops_template) - 1:
        if str(ops_template.iat[i, 7]) != 'nan':
            sheet3.write_merge(i + 1, i + 2, 7, 7, ops_template.iat[i, 7])
        else:
            sheet3.write_merge(i + 1, i + 2, 7, 7, "")
        i = i + 2

    i = 0
    while i <= len(ops_template) - 1:
        if str(ops_template.iat[i, 8]) != 'nan':
            sheet3.write_merge(i + 1, i + 2, 8, 8, ops_template.iat[i, 8])
        else:
            sheet3.write_merge(i + 1, i + 2, 8, 8, "")
        i = i + 2

    print("genNewOSC: XLWT完成光路编码、理论光路长度、理论衰耗的合并")

    # 保存最终文件并结束计时
    excelCopy.save('genNewNewOSC.xls')
    time_end = time.time()
    print("genNewOSC: DONE, Total Time Cost: ", time_end - time_start)
    alert(text="Targeted Excel Generated!\n用时："+str(time_end - time_start), title="处理结果", button="好的")


def wholeProcess():
    genNewOA()
    genNewOCP()
    genNewNewOSC()
    alert(text="三个分表格生成完毕！!\n请自行合并软件根目录下的三个文件", title="处理结果", button="好的")


if __name__ == "__main__":
    app = QtWidgets.QApplication([])

    widget = MyWidget()
    widget.resize(500, 600)
    widget.setWindowTitle("烽火系统表格自动生成工具 V1.0")
    widget.show()

    sys.exit(app.exec_())
