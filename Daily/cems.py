# -*- coding: utf-8 -*- 
# @Time : 2023/12/27 10:52 
# @Author : RoninX

import numpy as np
import re
import describetime

# 报警表常量
EVENTTYPE = "类型"
EQUIPMENTNAME = "设备名称"
FACTORY = "所属分厂"
PRODUCTIONPROCESS = "生产工序"
ALARMTYPE = "报警类型"
ALARMSTATUS = "报警状态"
ALARMDETAILS = "报警详情"
STARTTIME = "起始时间"
ALARMCAUSE = "报警原因"

# 定义各厂部常量
IRONMAKE = "炼铁厂"
STEELMAKE = "炼钢厂"
HOTROLL = "热轧厂"
COLDROLL = "冷轧厂"
SILICONSTEEL = "硅钢部"
STEELELECTRICITY = "钢电公司"
TRANSPORT = "运输部"
EAEP = "能环部"
WINSTEEL = "Winsteel"

# 定义事件类型的常量
CALIBRATION = "检校"
STOPPAGE = "启停"
OPERATECONDITION = "工况"
CEMSFAULT = "CEMS故障"

# 定义污染因子常量
SO2 = "二氧化硫"
SAND = "烟尘"
NOX = "氮氧化物"

# 定义其他常量
FACTORS = "污染因子"
TIME = "时间"

class Cems:
    # 处理检校类报警信息，返回日报的报警详情栏字符串
    def calibration(self,cemsarray,factory):
        self.cemsarray = cemsarray
        self.factory = factory
        calibrationarray = [[EQUIPMENTNAME,ALARMDETAILS,ALARMCAUSE]]
        for row in cemsarray[1:]:
            newrow = [row[1],row[4],row[5]]
            if(row[0] == CALIBRATION and row[2] == factory):
                calibrationarray = np.vstack((calibrationarray,newrow))
        calibrationstringarray = [[EQUIPMENTNAME,FACTORS,TIME]]
        pattern0 = re.compile(r'^.*?(?=\()')
        pattern1 = re.compile(r'\(([^)]*)\)')
        pattern2 = re.compile(r'\)(.*?)超')
        for row in calibrationarray[1:]:
            newrow = [pattern0.findall(row[1])[0],pattern1.search(row[1]).group(1),pattern2.findall(row[1])[0]]
            # print(newrow)
            calibrationstringarray = np.vstack((calibrationstringarray,newrow))

        calibrationstringarray_dispose = [[calibrationstringarray[-1][0]+calibrationstringarray[-1][1],calibrationstringarray[-1][2]]]
        if(len(calibrationstringarray)>2):
            for i in range(len(calibrationstringarray)-2,0,-1):
                newrow = [calibrationstringarray[i][0]+calibrationstringarray[i][1],calibrationstringarray[i][2]]
                calibrationstringarray_dispose = np.vstack((calibrationstringarray_dispose,newrow))
        # 这里是处理后每一组包括【点位因子，时间】的数组

        calibrationultimatestringarray = [] # 检校类报警最终的输出数组

        # 点位因子种类数组
        pointfactorstringarray = []
        for row in calibrationstringarray_dispose:
            pointfactorstringarray.append(row[0])
        pointfactorstringarray = list(set(pointfactorstringarray))

        temptime = []
        for i in pointfactorstringarray: # pointfactorstring与calibrationultimatestringarray每个相同位置元素是一一对应的
            for row in calibrationstringarray_dispose:
                if row[0] == i:
                    temptime.append(row[1])
            calibrationultimatestringarray.append(describetime.simplify_time_ranges(temptime))
            temptime = []

        if len(pointfactorstringarray) == 1:
            return  calibrationultimatestringarray[0]+ "，" + pointfactorstringarray[0] + "超限"
        # 将时间与事件的字符串拼接成最终结果
        ultimatereturnstring = ""
        ultimatereturncount = 0
        for sss in pointfactorstringarray:
            if ultimatereturnstring != "":
                ultimatereturnstring += "\n"
            ultimatereturnstring += str(ultimatereturncount+1) + "." + calibrationultimatestringarray[ultimatereturncount]+ "，" + pointfactorstringarray[ultimatereturncount] + "超限"
            ultimatereturncount += 1
        return ultimatereturnstring

    def stoppage(self,cemsarray,factory):
        self.cemsarray = cemsarray
        self.factory = factory
        stoppagearray = [[EQUIPMENTNAME, ALARMDETAILS, ALARMCAUSE]]
        for row in cemsarray[1:]:
            newrow = [row[1], row[4], row[5]]
            if (row[0] == STOPPAGE and row[2] == factory):
                stoppagearray = np.vstack((stoppagearray, newrow))
        stoppagestringarray = [[EQUIPMENTNAME, FACTORS, TIME]]
        pattern0 = re.compile(r'^.*?(?=\()')
        pattern1 = re.compile(r'\(([^)]*)\)')
        pattern2 = re.compile(r'\)(.*?)超')
        for row in stoppagearray[1:]:
            newrow = [pattern0.findall(row[1])[0], pattern1.search(row[1]).group(1), pattern2.findall(row[1])[0]]
            stoppagestringarray = np.vstack((stoppagestringarray, newrow))

        stoppagestringarray_dispose = [[stoppagestringarray[-1][0] + stoppagestringarray[-1][1],stoppagestringarray[-1][2]]]

        if (len(stoppagestringarray) > 2):
            for i in range(len(stoppagestringarray) - 2, 0, -1):
                newrow = [stoppagestringarray[i][0] + stoppagestringarray[i][1], stoppagestringarray[i][2]]
                stoppagestringarray_dispose = np.vstack((stoppagestringarray_dispose, newrow))

        stoppageultimatestringarray = []  # 启停类报警最终的输出数组

        # 点位因子种类数组
        pointfactorstringarray = []
        for row in stoppagestringarray_dispose:
            pointfactorstringarray.append(row[0])
        pointfactorstringarray = list(set(pointfactorstringarray))

        temptime = []
        for i in pointfactorstringarray:  # pointfactorstring与calibrationultimatestringarray每个相同位置元素是一一对应的
            for row in stoppagestringarray_dispose:
                if row[0] == i:
                    temptime.append(row[1])
            stoppageultimatestringarray.append(describetime.simplify_time_ranges(temptime))
            temptime = []

        if len(pointfactorstringarray) == 1:
            return stoppageultimatestringarray[0] + "，" + pointfactorstringarray[0] + "超限"
        # 将时间与事件的字符串拼接成最终结果
        ultimatereturnstring = ""
        ultimatereturncount = 0
        for sss in pointfactorstringarray:
            if ultimatereturnstring != "":
                ultimatereturnstring += "\n"
            ultimatereturnstring += str(ultimatereturncount + 1) + "." + stoppageultimatestringarray[ultimatereturncount] + "，" + pointfactorstringarray[ultimatereturncount] + "超限"
            ultimatereturncount += 1
        return ultimatereturnstring

    def operatecondition(self,cemsarray,factory):
        self.cemsarray = cemsarray
        self.factory = factory
        operateconditionarray = [[EQUIPMENTNAME, ALARMDETAILS, ALARMCAUSE]]
        for row in cemsarray[1:]:
            newrow = [row[1], row[4], row[5]]
            if (row[0] == OPERATECONDITION and row[2] == factory):
                operateconditionarray = np.vstack((operateconditionarray,newrow))
        operateconditionstringarray = [[EQUIPMENTNAME, FACTORS, TIME, ALARMCAUSE]]
        pattern0 = re.compile(r'^.*?(?=\()')
        pattern1 = re.compile(r'\(([^)]*)\)')
        pattern2 = re.compile(r'\)(.*?)超')
        for rerow in operateconditionarray[1:]:
            renewrow = [pattern0.findall(rerow[1])[0], pattern1.search(rerow[1]).group(1),pattern2.findall(rerow[1])[0], rerow[2]]
            operateconditionstringarray = np.vstack((operateconditionstringarray, renewrow))

        operateconditionstringarray_dispose = [[operateconditionstringarray[-1][0] + operateconditionstringarray[-1][1] + operateconditionstringarray[-1][3],operateconditionstringarray[-1][0],operateconditionstringarray[-1][1],operateconditionstringarray[-1][2],operateconditionstringarray[-1][3]]]

        if len(operateconditionstringarray) > 2:
            for i in range(len(operateconditionstringarray) - 2, 0, -1):
                rerenewrow = [operateconditionstringarray[i][0] + operateconditionstringarray[i][1] + operateconditionstringarray[i][3],operateconditionstringarray[i][0],operateconditionstringarray[i][1],operateconditionstringarray[i][2],operateconditionstringarray[i][3]]
                operateconditionstringarray_dispose = np.vstack((operateconditionstringarray_dispose,rerenewrow))

        uniquestringarray = []
        for rerererow in operateconditionstringarray_dispose:
            uniquestringarray.append(rerererow[0])
        uniquestringarray = list(set(uniquestringarray))

        operateconditionsimpletimearray =[]
        uniquetimearray = []
        uniquepointfactor = []
        uniquealarmcause = []
        for rererererow in uniquestringarray:
            for rowa in operateconditionstringarray_dispose:
                if rowa[0] == rererererow:
                    uniquetimearray.append(rowa[3])
                    pointfactor = rowa[1] + rowa[2]
                    alarmcause = rowa[4]
            operateconditionsimpletimearray.append(describetime.simplify_time_ranges(uniquetimearray))
            uniquepointfactor.append(pointfactor)
            uniquealarmcause.append(alarmcause)
            uniquetimearray = []

        if len(operateconditionsimpletimearray) == 1:
            return operateconditionsimpletimearray[0] + "，" + operateconditionstringarray_dispose[0][1] + operateconditionstringarray_dispose[0][2] + "超限，原因：" + operateconditionstringarray_dispose[0][4]

        ultimatereturnstring = ""
        ultimatereturncount = 0
        for rowb in uniquestringarray:
            if ultimatereturnstring != "":
                ultimatereturnstring += "\n"
            ultimatereturnstring += str(ultimatereturncount + 1) + "." + operateconditionsimpletimearray[ultimatereturncount] + "，" + uniquepointfactor[ultimatereturncount] + "超限，原因：" + uniquealarmcause[ultimatereturncount]
            ultimatereturncount += 1
        return ultimatereturnstring

    def cemsfault(self,cemsarray,factory):
        self.cemsarray = cemsarray
        self.factory = factory
        cemsfaultarray = [[EQUIPMENTNAME, ALARMDETAILS, ALARMCAUSE]]
        for row in cemsarray[1:]:
            newrow = [row[1],row[4],row[5]]
            if row[0] == CEMSFAULT and row[2] == factory:
                cemsfaultarray = np.vstack((cemsfaultarray,newrow))
        cemsfaultstringarray = [[EQUIPMENTNAME, FACTORS, TIME, ALARMCAUSE]]
        pattern0 = re.compile(r'^.*?(?=\()')
        pattern1 = re.compile(r'\(([^)]*)\)')
        pattern2 = re.compile(r'\)(.*?)超')
        for rerow in cemsfaultarray[1:]:
            renewrow = [pattern0.findall(rerow[1])[0], pattern1.search(rerow[1]).group(1),pattern2.findall(rerow[1])[0], rerow[2]]
            cemsfaultstringarray = np.vstack((cemsfaultstringarray, renewrow))

        cemsfaultstringarray_dispose = [[cemsfaultstringarray[-1][0] + cemsfaultstringarray[-1][1] + cemsfaultstringarray[-1][3],cemsfaultstringarray[-1][0],cemsfaultstringarray[-1][1],cemsfaultstringarray[-1][2],cemsfaultstringarray[-1][3]]]
        if len(cemsfaultstringarray) > 2:
            for i in range(len(cemsfaultstringarray) - 2, 0, -1):
                rerenewrow = [cemsfaultstringarray[i][0] + cemsfaultstringarray[i][1] + cemsfaultstringarray[i][3],cemsfaultstringarray[i][0],cemsfaultstringarray[i][1],cemsfaultstringarray[i][2],cemsfaultstringarray[i][3]]
                cemsfaultstringarray_dispose = np.vstack((cemsfaultstringarray_dispose,rerenewrow))

        uniquestringarray = []
        for rererenewrow in cemsfaultstringarray_dispose:
            uniquestringarray.append(rererenewrow[0])
        uniquestringarray = list(set(uniquestringarray))

        cemsfaultsimpletimearray = []
        uniquetimearray = []
        uniquepointfactor = []
        uniquealarmcause = []
        for rerererenewrow in uniquestringarray:
            for rowa in cemsfaultstringarray_dispose:
                if rowa[0] == rerererenewrow:
                    uniquetimearray.append(rowa[3])
                    pointfactor = rowa[1] + rowa[2]
                    alarmcause = rowa[4]
            cemsfaultsimpletimearray.append(describetime.simplify_time_ranges(uniquetimearray))
            uniquepointfactor.append(pointfactor)
            uniquealarmcause.append(alarmcause)
            uniquetimearray = []

        if len(cemsfaultsimpletimearray) == 1:
            return cemsfaultsimpletimearray[0] + "，" + cemsfaultstringarray_dispose[0][1] + cemsfaultstringarray_dispose[0][2] + "超限，原因：" + cemsfaultstringarray_dispose[0][4]

        ultimatereturnstring = ""
        ultimatereturncount = 0
        for rowb in uniquestringarray:
            if ultimatereturnstring != "":
                ultimatereturnstring += "\n"
            ultimatereturnstring += str(ultimatereturncount + 1) + "." + cemsfaultsimpletimearray[ultimatereturncount] + "，" + uniquepointfactor[ultimatereturncount] + "超限，原因：" + uniquealarmcause[ultimatereturncount]
            ultimatereturncount += 1
        return ultimatereturnstring



