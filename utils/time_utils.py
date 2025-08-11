from datetime import datetime
import re

def get_minimum_distance(letter):
    """计算冒号之间的最小距离"""
    position = []
    for i in range(len(letter)):
        if letter[i] == ':':
            position.append(i)

    if len(position) > 1:
        j = len(position) - 1
        distance = []
        while j >= 1:
            d = position[j] - position[j - 1]
            j = j - 1
            distance.append(d)
        return min(distance)
    else:
        return None

def daily_working_time(time_list_normalized):
    """计算每日工作时间"""
    time_part = []
    i = len(time_list_normalized)

    while (i >= 1):
        time_difference = (time_list_normalized[i - 1] - time_list_normalized[i - 2]).total_seconds() / (60 * 60)
        i = i - 2
        time_part.append(time_difference)

    return round(sum(time_part), 2)

def parse_time_string(raw_time_str):
    """解析时间字符串，提取有效的时间"""
    time_list = []
    if '\n' in raw_time_str:
        time_list = raw_time_str.split('\n')
    else:
        time_pattern = r'\d{1,2}:\d{2}'
        time_list = re.findall(time_pattern, raw_time_str)

    return [t.strip() for t in time_list if t.strip()]

def validate_time_format(time_str):
    """验证时间格式是否有效"""
    try:
        if ':' in time_str and len(time_str.split(':')) == 2:
            hour, minute = time_str.split(':')
            if 0 <= int(hour) <= 23 and 0 <= int(minute) <= 59:
                return True
        return False
    except:
        return False

def normalize_time_list(time_list):
    """规范化时间列表，返回datetime对象列表"""
    time_list_normalized = []
    for time_str in time_list:
        try:
            if validate_time_format(time_str):
                date_time_obj = datetime.strptime(time_str, '%H:%M')
                time_list_normalized.append(date_time_obj)
        except ValueError:
            continue
    return time_list_normalized 