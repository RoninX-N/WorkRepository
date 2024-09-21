import numpy as np
from datetime import datetime

# 创建示例数组，第一行是表头，时间列为字符串类型
data = np.array([
    ['id', 'col1', 'col2', 'date'],
    [1, 2, 3, '2023-01-01 12:00'],
    [4, 5, 6, '2024-01-01 12:00'],
    [7, 8, 9, '2022-01-01 12:00'],
    [10, 11, 12, '2025-01-01 12:00']
], dtype=object)

print(type(data))

# 指定的时间点
specified_time = datetime(2023, 6, 1, 0, 0)

# 分离表头和数据
header = data[0]
data = data[1:]

# 将日期字符串转换为datetime对象
dates = np.array([datetime.strptime(date, "%Y-%m-%d %H:%M") for date in data[:, 3]])
print(dates)

# 使用布尔索引来过滤掉日期早于指定时间点的行
filtered_data = data[dates >= specified_time]

# 重新添加表头
filtered_data_with_header = np.vstack([header, filtered_data])

print(filtered_data_with_header)



