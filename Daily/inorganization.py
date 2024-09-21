# -*- coding: utf-8 -*- 
# @Time : 2024/1/8 12:25 
# @Author : RoninX

from pprint import pprint

import pandas as pd
import numpy as np
from datetime import datetime
import describetime

# 报警表常量
EVENTTYPE = "类型"
EQUIPMENTNAME = "设备名称"
EQUIPMENTTYPE = "设备类型"
FACTORY = "所属分厂"
PRODUCTIONPROCESS = "生产工序"
ALARMTYPE = "报警类型"
ALARMSTATUS = "报警状态"
ALARMDETAILS = "报警详情"
STARTTIME = "起始时间"
ALARMCAUSE = "建议排查原因"

# 定义各厂部常量
IRONMAKE = "炼铁厂"
STEELMAKE = "炼钢厂"
HOTROLL = "热轧厂"
COLDROLL = "冷轧厂"
SILICONSTEEL = "硅钢部"
STEELELECTRICITY = "钢电公司"
TRANSPORT = "运输部"
EAEP = "能环部"

# 定义其他常量
FACTORS = "污染因子"
TIME = "时间"
GOVERNANCEFACILITIES = "治理设施"
TVOC = "TVOC监测仪"

UNLINKAGE = "未联动开启"
DIFFERENTIALPRESSUREOVERRUN = "除尘器连续超出额定压差范围"
FANSCURRENTOVERRUN = "风机电流连续超过额定电流"

# 调用方法返回对应字符串
class Inorganization:
    def governance(self,inorganizationarray,factory):
        self.inorganizationarray = inorganizationarray
        self.factory = factory

        # 分未联动开启、风机电流连续超过额定电流、除尘器连续超出额定压差范围 三种来处理
        differentialpressureoverrun = [[inorganizationarray[0][0]+inorganizationarray[0][7],EQUIPMENTNAME,EQUIPMENTTYPE,FACTORY,ALARMTYPE,ALARMSTATUS,ALARMDETAILS,STARTTIME,ALARMCAUSE]]
        fanscurrentoverrun = [[inorganizationarray[0][0]+inorganizationarray[0][7],EQUIPMENTNAME,EQUIPMENTTYPE,FACTORY,ALARMTYPE,ALARMSTATUS,ALARMDETAILS,STARTTIME,ALARMCAUSE]]
        unlinkage = [[inorganizationarray[0][5]+inorganizationarray[0][7],EQUIPMENTNAME,EQUIPMENTTYPE,FACTORY,ALARMTYPE,ALARMSTATUS,ALARMDETAILS,STARTTIME,ALARMCAUSE]]
        for row in inorganizationarray[1:]:
            if row[2] == factory and row[3] == UNLINKAGE and row[1] == GOVERNANCEFACILITIES:
                newrow = [row[5] + row[7], row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]]
                unlinkage = np.vstack((unlinkage,newrow))
            elif row[2] == factory and row[3] == DIFFERENTIALPRESSUREOVERRUN and row[1] == GOVERNANCEFACILITIES:
                newrow = [row[0] + row[7], row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]]
                differentialpressureoverrun = np.vstack((differentialpressureoverrun,newrow))
            elif row[2] == factory and row[3] == FANSCURRENTOVERRUN and row[1] == GOVERNANCEFACILITIES:
                newrow = [row[0] + row[7], row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]]
                fanscurrentoverrun = np.vstack((fanscurrentoverrun, newrow))

        # 求未联动开启唯一数组
        unlinkageuniquearray = []
        for row in unlinkage[1:]:
            unlinkageuniquearray.append(row[0])
        unlinkageuniquearray = list(set(unlinkageuniquearray))

        # 求压差超限唯一数组
        differentialpressureoverrununiquearray = []
        for row in differentialpressureoverrun[1:]:
            differentialpressureoverrununiquearray.append(row[0])
        differentialpressureoverrununiquearray = list(set(differentialpressureoverrununiquearray))

        # 求风机电流超限唯一数组
        fanscurrentoverrununiquearray = []
        for row in fanscurrentoverrun[1:]:
            fanscurrentoverrununiquearray.append(row[0])
        fanscurrentoverrununiquearray = list(set(fanscurrentoverrununiquearray))

        # @终极之爆炸阿姆斯特朗回旋兼拉格朗日中值定理据切比雪夫不等式与一身的无组织环保设备报警字符串
        inorganiseultimatestring = ""

        # 对无组织输出string的辅助变量
        count = 1

        # 未联动开启的字符串处理和添加
        unlinkagetemptime = []
        unlinkagedetailstring = ""
        unlinkagecausestring = ""
        for elementi in unlinkageuniquearray:
            for elementj in unlinkage[1:]:
                if elementj[0] == elementi:
                    unlinkagedetailstring = elementj[6]
                    unlinkagecausestring = elementj[8]
                    if isinstance(elementj[7],datetime):
                        unlinkagetemptime.append(elementj[7].strftime("%Y-%m-%d %H:%M:%S"))
                    else:
                        unlinkagetemptime.append(elementj[7])
            # 一系列操作
            # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            # 将时间排序
            datetime_objects = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in unlinkagetemptime])
            if inorganiseultimatestring == "":
                inorganiseultimatestring += str(count) + "."
            else:
                inorganiseultimatestring += "\n" + str(count) + "."
            count = count + 1
            tempday = datetime_objects[0]
            inorganiseultimatestring += f"{datetime_objects[0].strftime('%Y年%m月%d日%H:%M')}"
            if len(datetime_objects) > 1:
                for time in datetime_objects[1:]:
                    if tempday.strftime('%Y年%m月%d日') == time.strftime('%Y年%m月%d日'):
                        inorganiseultimatestring += f"、{time.strftime('%H:%M')}"
                    else:
                        inorganiseultimatestring += f"，{time.strftime('%d日%H:%M')}"
                        tempday = time
            unlinkagedetailstring = unlinkagedetailstring.replace("未同步开启","未联动开启")
            inorganiseultimatestring += "，" + unlinkagedetailstring + "，原因：" + unlinkagecausestring
            # 归空
            unlinkagetemptime = []
            # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        # 压差超限的字符串处理和添加
        differentialpressureoverruntemptime = []
        differentialpressureoverrunequipnamestring = ""
        differentialpressureoverruncausestring = ""
        for elementi in differentialpressureoverrununiquearray:
            for elementj in differentialpressureoverrun[1:]:
                if elementj[0] == elementi:
                    differentialpressureoverrunequipnamestring = elementj[1]
                    differentialpressureoverruncausestring = elementj[8]
                    if isinstance(elementj[7],datetime):
                        differentialpressureoverruntemptime.append(elementj[7].strftime("%Y-%m-%d %H:%M:%S"))
                    else:
                        differentialpressureoverruntemptime.append(elementj[7])
            datetime_objects = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in differentialpressureoverruntemptime])
            if inorganiseultimatestring == "":
                inorganiseultimatestring += str(count) + "."
            else:
                inorganiseultimatestring += "\n" + str(count) + "."
            count = count + 1
            tempday = datetime_objects[0]
            inorganiseultimatestring += f"{datetime_objects[0].strftime('%Y年%m月%d日%H:%M')}"
            if len(datetime_objects) > 1:
                for time in datetime_objects[1:]:
                    if tempday.strftime('%Y年%m月%d日') == time.strftime('%Y年%m月%d日'):
                        inorganiseultimatestring += f"、{time.strftime('%H:%M')}"
                    else:
                        inorganiseultimatestring += f"，{time.strftime('%d日%H:%M')}"
                        tempday = time
            inorganiseultimatestring += "，" + differentialpressureoverrunequipnamestring + "压差超限，原因：" + differentialpressureoverruncausestring
            differentialpressureoverruntemptime = []

        # 风机电流超限字符串处理和添加
        fanscurrentoverruntemptime = []
        fanscurrentoverrunequipnamestring = ""
        fanscurrentoverruncausestring = ""
        for elementi in fanscurrentoverrununiquearray:
            for elementj in fanscurrentoverrun[1:]:
                if elementj[0] == elementi:
                    fanscurrentoverrunequipnamestring = elementj[1]
                    fanscurrentoverruncausestring = elementj[8]
                    if isinstance(elementj[7],datetime):
                        fanscurrentoverruntemptime.append(elementj[7].strftime("%Y-%m-%d %H:%M:%S"))
                    else:
                        fanscurrentoverruntemptime.append(elementj[7])
            datetime_objects = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in fanscurrentoverruntemptime])
            if inorganiseultimatestring == "":
                inorganiseultimatestring += str(count) + "."
            else:
                inorganiseultimatestring += "\n" + str(count) + "."
            count = count + 1
            tempday = datetime_objects[0]
            inorganiseultimatestring += f"{datetime_objects[0].strftime('%Y年%m月%d日%H:%M')}"
            if len(datetime_objects) > 1:
                for time in datetime_objects[1:]:
                    if tempday.strftime('%Y年%m月%d日') == time.strftime('%Y年%m月%d日'):
                        inorganiseultimatestring += f"、{time.strftime('%H:%M')}"
                    else:
                        inorganiseultimatestring += f"，{time.strftime('%d日%H:%M')}"
                        tempday = time
            inorganiseultimatestring += "，" + fanscurrentoverrunequipnamestring + "风机电流超限，原因：" + fanscurrentoverruncausestring
            fanscurrentoverruntemptime = []

        return inorganiseultimatestring

    def TSP(self,inorganizationarray,factory):
        self.inorganizationarray = inorganizationarray
        self.factory = factory
        TSPultimatestring = ""
        TSPstringcount = 1
        TSParray = [[EQUIPMENTNAME,ALARMTYPE,STARTTIME]]
        TSPtemptime = []
        for row in inorganizationarray[1:]:
            if row[2] == factory and row[1] == "TSP":
                newrow = [row[0],row[3],row[6]]
                TSParray = np.vstack((TSParray,newrow))
        if len(TSParray) == 2:
            if isinstance(TSParray[1][2],datetime):
                TSPtemptime.append(TSParray[1][2].strftime("%Y-%m-%d %H:%M:%S"))
            else:
                TSPtemptime.append(TSParray[1][2])
            datetime_objects = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in TSPtemptime])
            return datetime_objects[0].strftime('%Y年%m月%d日%H:%M') + "，" + TSParray[1][0] + TSParray[1][1]
        TSPuniqueequipmentarray = []
        for row in TSParray[1:]:
            TSPuniqueequipmentarray.append(row[0])
        TSPuniqueequipmentarray = list(set(TSPuniqueequipmentarray))
        tempequipmentname = ""
        tempalarmtype = ""
        for x in TSPuniqueequipmentarray:
            for temprow in TSParray[1:]:
                if temprow[0] == x:
                    if isinstance(temprow[2],datetime):
                        TSPtemptime.append(temprow[2].strftime("%Y-%m-%d %H:%M:%S"))
                    else:
                        TSPtemptime.append(temprow[2])
                    tempequipmentname = temprow[0]
                    tempalarmtype = temprow[1]
            TSPtemptime = [str(time) if isinstance(time, datetime) else time for time in TSPtemptime]
            TSPtemptime = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in TSPtemptime])
            TSPultimatestring += str(TSPstringcount) + "."
            TSPstringcount += 1
            tempday = TSPtemptime[0]
            TSPultimatestring += f"{TSPtemptime[0].strftime('%Y年%m月%d日%H:%M')}"
            for time in TSPtemptime[1:]:
                if tempday.strftime('%Y年%m月%d日') == time.strftime('%Y年%m月%d日'):
                    TSPultimatestring += f"、{time.strftime('%H:%M')}"
                else:
                    TSPultimatestring += f"，{time.strftime('%d日%H:%M')}"
                    tempday = time
            TSPultimatestring += "，" + tempequipmentname + tempalarmtype
            if x != TSPuniqueequipmentarray[-1]:
                TSPultimatestring += "\n"
            else:
                None
        return TSPultimatestring

    def TVOC(self,inorganizationarray,factory):
        self.inorganizationarray = inorganizationarray
        self.factory = factory
        TVOCultimatestring = ""
        TVOCstringcount = 1
        TVOCarray = [[EQUIPMENTNAME, ALARMTYPE, STARTTIME]]
        TVOCtemptime = []
        for row in inorganizationarray[1:]:
            if row[2] == factory and row[1] == TVOC:
                newrow = [row[0],row[5],row[6]]
                TVOCarray = np.vstack((TVOCarray,newrow))
        if len(TVOCarray) == 2:
            if isinstance(TVOCarray[1][2],datetime):
                TVOCtemptime.append(TVOCarray[1][2].strftime("%Y-%m-%d %H:%M:%S"))
            else:
                TVOCtemptime.append(TVOCarray[1][2])
            datetime_objects = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in TVOCtemptime])
            return datetime_objects[0].strftime('%Y年%m月%d日%H:%M') + "，" + TVOCarray[1][0] + TVOCarray[1][1]
        TVOCuniqueequipmentarray = []
        for row in TVOCarray[1:]:
            TVOCuniqueequipmentarray.append(row[0])
        TVOCuniqueequipmentarray = list(set(TVOCuniqueequipmentarray))
        tempequipmentname = ""
        tempalarmdetail = ""
        for row in TVOCuniqueequipmentarray:
            for temprow in TVOCarray[1:]:
                if temprow[0] == row:
                    if isinstance(temprow[0],datetime):
                        TVOCtemptime.append(temprow[2].strftime("%Y-%m-%d %H:%M:%S"))
                    else:
                        TVOCtemptime.append(temprow[2])
                    tempequipmentname = temprow[0]
                    tempalarmdetail = temprow[1]
            TVOCtemptime = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S') for time_str in TVOCtemptime])
            TVOCultimatestring += str(TVOCstringcount) + "."
            TVOCstringcount += 1
            tempday = TVOCtemptime[0]
            TVOCultimatestring += f"{TVOCtemptime[0].strftime('%Y年%m月%d日%H:%M')}"
            for time in TVOCtemptime[1:]:
                if tempday.strftime('%Y年%m月%d日') == time.strftime('%Y年%m月%d日'):
                    TVOCultimatestring += f"、{time.strftime('%H:%M')}"
                else:
                    TVOCultimatestring += f"，{time.strftime('%d日%H:%M')}"
                    tempday = time
            TVOCultimatestring += "，" + tempequipmentname +tempalarmdetail
        return TVOCultimatestring

# 调试用

def main():
    np.set_printoptions(threshold=np.inf)
    sheet = pd.read_excel("D:\\elseFile\\work\\Codeworkbench\\ChartTemplate\\报警历史（到2024.1.10下午4点）.xlsx", header=None)
    inorganizationarray = [
        [EQUIPMENTNAME, EQUIPMENTTYPE, FACTORY, ALARMTYPE, ALARMSTATUS, ALARMDETAILS, STARTTIME, ALARMCAUSE]]
    equipmentnamecolumn = 0
    equipmenttypecolumn = 0
    factorycolumn = 0
    alarmtypecolumn = 0
    alarmstatuscolumn = 0
    alarmdetailscolumn = 0
    starttimecolumn = 0
    alarmcausecolumn = 0

    columncout = 0
    first_row = sheet.iloc[0]
    for count in first_row:
        if count == EQUIPMENTNAME:
            equipmentnamecolumn = columncout
        if count == EQUIPMENTTYPE:
            equipmenttypecolumn = columncout
        if count == FACTORY:
            factorycolumn = columncout
        if count == ALARMTYPE:
            alarmtypecolumn = columncout
        if count == ALARMSTATUS:
            alarmstatuscolumn = columncout
        if count == ALARMDETAILS:
            alarmdetailscolumn = columncout
        if count == STARTTIME:
            starttimecolumn = columncout
        if count == ALARMCAUSE:
            alarmcausecolumn = columncout
        columncout = columncout + 1

    rowcount = 1
    for row in sheet.iloc[1:].itertuples(index=False):
        if row[equipmenttypecolumn] == GOVERNANCEFACILITIES or row[equipmenttypecolumn] == "TSP" or row[equipmenttypecolumn] == TVOC:
            if pd.isnull(sheet.at[rowcount,alarmcausecolumn]):
                newrow = [row[equipmentnamecolumn],row[equipmenttypecolumn],row[factorycolumn],row[alarmtypecolumn],row[alarmstatuscolumn],row[alarmdetailscolumn],row[starttimecolumn],"待反馈"]
            else:
                newrow = [row[equipmentnamecolumn],row[equipmenttypecolumn],row[factorycolumn],row[alarmtypecolumn],row[alarmstatuscolumn],row[alarmdetailscolumn],row[starttimecolumn],row[alarmcausecolumn]]
            inorganizationarray = np.vstack((inorganizationarray, newrow))
        rowcount = rowcount + 1
    inorganization_instance = Inorganization()
    inorganization_instance.governance(inorganizationarray,STEELMAKE)

if __name__ == "__main__":
    main()
