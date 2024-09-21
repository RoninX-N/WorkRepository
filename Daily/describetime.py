# -*- coding: utf-8 -*- 
# @Time : 2024/1/2 15:17 
# @Author : RoninX

from datetime import datetime, timedelta
from dateutil.parser import parse

def simplify_time_ranges(time_strings):
    # 将时间字符串转换为 datetime 对象，并按时间排序
    datetime_objects = sorted([datetime.strptime(time_str, '%Y-%m-%d %H:%M') for time_str in time_strings])
    # 初始化结果字符串
    result_string = ""
    continuousflag = 0
    for i in range(len(datetime_objects)):
        if (i == 0):
            result_string += f"{datetime_objects[i].strftime('%Y年%m月%d日%H:%M')}"
            tempday = datetime_objects[i]
            # 保证输出
            if len(datetime_objects) > 1:
                endtime = datetime_objects[0]
            else:
                break
        else:
            # 判断是否连续
            if datetime_objects[i] == endtime + timedelta(hours=1):
                endtime = datetime_objects[i]
                # 避免重复输出的值
                continuousflag = 1
                # 保证输出
                if i + 1 > len(datetime_objects) - 1:
                    if tempday.strftime('%Y年%m月%d日') == endtime.strftime('%Y年%m月%d日'):
                        result_string += f"-{endtime.strftime('%H:%M')}"
                    else:
                        result_string += f"-{endtime.strftime('%d日%H:%M')}"
                    break
            else:  # 不连续的情况
                if continuousflag == 1:
                    # 这里加if判断是否到第二天
                    if tempday.strftime('%Y年%m月%d日') == endtime.strftime('%Y年%m月%d日'):
                        result_string += f"-{endtime.strftime('%H:%M')}"
                    else:
                        result_string += f"-{endtime.strftime('%d日%H:%M')}"
                        tempday = datetime_objects[i]
                    continuousflag = 0
                # 这里加if判断是否去掉日
                if tempday.strftime('%Y年%m月%d日') == datetime_objects[i].strftime('%Y年%m月%d日'):
                    result_string += f"、{datetime_objects[i].strftime('%H:%M')}"
                else:
                    result_string += f"，{datetime_objects[i].strftime('%m月%d日%H:%M')}"
                    tempday = datetime_objects[i]
                if i + 1 > len(datetime_objects) - 1:
                    break;
                else:
                    endtime = datetime_objects[i]
    return result_string