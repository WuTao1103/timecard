from datetime import datetime
import re
import logging


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
    """
    增强的时间字符串解析，支持多种分隔符和格式
    支持格式:
    - 换行分隔: "10:36\n11:18\n11:33\n21:10"
    - 空格分隔: "10:36  11:18 11:33 21:10"
    - 逗号分隔: "10:36,11:18,11:33,21:10"
    - 制表符分隔: "10:36\t11:18\t11:33\t21:10"
    - 混合分隔符
    """
    import re
    
    if not raw_time_str or str(raw_time_str).strip() in ['nan', '', 'None']:
        return []

    raw_time_str = str(raw_time_str).strip()
    time_list = []

    # 方法1: 处理换行分隔
    if '\n' in raw_time_str:
        # 先按换行分割，然后处理每一行
        lines = raw_time_str.split('\n')
        print(f"📝 换行分割: {raw_time_str} -> {lines}")
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 如果行内还有空格分隔的时间，进一步分割
            if ' ' in line:
                # 使用正则表达式提取该行中的所有时间
                time_pattern = r'\b\d{1,2}:\d{2}\b'
                line_times = re.findall(time_pattern, line)
                time_list.extend(line_times)
                print(f"📝 行内空格分割: '{line}' -> {line_times}")
            else:
                # 单行单个时间
                if re.match(r'^\d{1,2}:\d{2}$', line):
                    time_list.append(line)
                    print(f"📝 单行时间: '{line}'")

    # 方法2: 处理其他分隔符（空格、制表符、逗号等）
    elif any(sep in raw_time_str for sep in [' ', '\t', ',', ';']):
        # 先用正则表达式提取所有可能的时间
        time_pattern = r'\b\d{1,2}:\d{2}\b'
        time_matches = re.findall(time_pattern, raw_time_str)
        if time_matches:
            time_list = time_matches
            print(f"📝 正则提取（有分隔符）: {raw_time_str} -> {time_list}")
        else:
            # 如果正则没匹配到，尝试分割
            # 尝试多种分隔符
            for separator in [' ', '\t', ',', ';', '  ', '   ']:
                if separator in raw_time_str:
                    time_list = raw_time_str.split(separator)
                    break
            print(f"📝 分隔符分割: {raw_time_str} -> {time_list}")

    # 方法3: 使用正则表达式提取所有时间格式（兜底方案）
    else:
        time_pattern = r'\b\d{1,2}:\d{2}\b'
        time_list = re.findall(time_pattern, raw_time_str)
        print(f"📝 正则提取（无分隔符）: {raw_time_str} -> {time_list}")

    # 清理和验证时间列表
    cleaned_times = []
    for time_str in time_list:
        time_str = time_str.strip()
        if time_str and time_str != '' and ':' in time_str:
            # 基础格式验证
            if re.match(r'^\d{1,2}:\d{2}$', time_str):
                cleaned_times.append(time_str)
            else:
                print(f"⚠️ 跳过无效格式: '{time_str}'")

    print(f"🔄 最终清理后: {cleaned_times}")
    return cleaned_times


def validate_time_format(time_str):
    """增强的时间格式验证"""
    if not time_str or not isinstance(time_str, str):
        return False

    time_str = time_str.strip()

    # 检查基本格式
    if not re.match(r'^\d{1,2}:\d{2}$', time_str):
        return False

    try:
        if ':' in time_str and len(time_str.split(':')) == 2:
            hour, minute = time_str.split(':')
            hour, minute = int(hour), int(minute)

            # 验证小时和分钟的有效性
            if 0 <= hour <= 23 and 0 <= minute <= 59:
                return True
        return False
    except (ValueError, AttributeError):
        return False


def normalize_time_list(time_list):
    """
    规范化时间列表，返回datetime对象列表
    增加了更详细的错误处理和日志记录
    """
    if not time_list:
        return []

    time_list_normalized = []
    invalid_times = []

    for time_str in time_list:
        time_str = str(time_str).strip()

        if not time_str or time_str in ['', 'nan', 'None']:
            continue

        if validate_time_format(time_str):
            try:
                date_time_obj = datetime.strptime(time_str, '%H:%M')
                time_list_normalized.append(date_time_obj)
            except ValueError as e:
                invalid_times.append(f"{time_str} (解析错误: {e})")
                print(f"⚠️ 时间解析失败: {time_str} - {e}")
        else:
            invalid_times.append(f"{time_str} (格式无效)")
            print(f"⚠️ 时间格式无效: {time_str}")

    if invalid_times:
        print(f"🚨 发现 {len(invalid_times)} 个无效时间: {invalid_times}")

    print(f"✅ 成功解析 {len(time_list_normalized)} 个有效时间")
    return time_list_normalized


def detect_time_anomalies(raw_time_str, employee_name, column_idx):
    """
    检测时间数据异常
    返回异常类型和详细信息，包含颜色映射
    """
    anomalies = []

    if not raw_time_str or str(raw_time_str).strip() in ['nan', '', 'None']:
        return anomalies

    raw_time_str = str(raw_time_str).strip()

    # 检测1: 冒号距离异常
    letter = [x for x in raw_time_str]
    min_distance = get_minimum_distance(letter)
    if min_distance == 3:
        anomalies.append({
            'type': 'colon_distance',
            'message': f'冒号距离异常 - 员工: {employee_name}, 列: {column_idx}, 最小距离: {min_distance}',
            'severity': 'warning',
            'color': 'FFC7CE',  # 浅红色
            'description': '时间格式问题，冒号前后数字位数异常'
        })

    # 检测2: 解析时间
    time_list = parse_time_string(raw_time_str)
    valid_times = normalize_time_list(time_list)

    # 检测3: 奇数时间记录
    if len(valid_times) % 2 != 0:
        anomalies.append({
            'type': 'odd_time_count',
            'message': f'奇数时间记录 - 员工: {employee_name}, 列: {column_idx}, 时间数量: {len(valid_times)}',
            'severity': 'error',
            'color': 'FF0000',  # 深红色
            'description': '打卡次数为奇数，无法配对计算工时'
        })

    # 检测4: 时间跨度异常
    if len(valid_times) >= 2:
        time_span = (valid_times[-1] - valid_times[0]).total_seconds() / 3600
        if time_span > 16:  # 工作时间跨度超过16小时
            anomalies.append({
                'type': 'long_work_span',
                'message': f'工作时间跨度异常 - 员工: {employee_name}, 列: {column_idx}, 跨度: {time_span:.1f}小时',
                'severity': 'warning',
                'color': 'FFD700',  # 金色
                'description': '单日工作时间跨度超过16小时，可能存在数据错误'
            })

    # 检测5: 时间顺序异常
    if len(valid_times) >= 2:
        for i in range(1, len(valid_times)):
            if valid_times[i] <= valid_times[i - 1]:
                anomalies.append({
                    'type': 'time_sequence_error',
                    'message': f'时间顺序异常 - 员工: {employee_name}, 列: {column_idx}, {valid_times[i - 1].strftime("%H:%M")} >= {valid_times[i].strftime("%H:%M")}',
                    'severity': 'error',
                    'color': 'FF8C00',  # 深橙色
                    'description': '打卡时间顺序混乱，后一个时间早于前一个时间'
                })
                break

    # 检测6: 时间格式无效
    invalid_times = []
    for time_str in time_list:
        if not validate_time_format(time_str):
            invalid_times.append(time_str)
    
    if invalid_times:
        anomalies.append({
            'type': 'invalid_time_format',
            'message': f'时间格式无效 - 员工: {employee_name}, 列: {column_idx}, 无效时间: {invalid_times}',
            'severity': 'error',
            'color': 'FF6B6B',  # 橙红色
            'description': '时间格式不符合HH:MM标准'
        })

    # 检测7: 解析错误
    if not time_list and raw_time_str not in ['nan', '', 'None']:
        anomalies.append({
            'type': 'parse_error',
            'message': f'解析错误 - 员工: {employee_name}, 列: {column_idx}, 原始数据: {raw_time_str}',
            'severity': 'error',
            'color': '9932CC',  # 紫色
            'description': '时间字符串无法正确解析'
        })

    # 检测8: 混合分隔符
    separators = ['\n', ' ', '\t', ',', ';']
    found_separators = [sep for sep in separators if sep in raw_time_str]
    if len(found_separators) > 1:
        anomalies.append({
            'type': 'mixed_separators',
            'message': f'混合分隔符 - 员工: {employee_name}, 列: {column_idx}, 分隔符: {found_separators}',
            'severity': 'warning',
            'color': '87CEEB',  # 天蓝色
            'description': '时间字符串包含多种分隔符，可能导致解析错误'
        })

    return anomalies


def format_time_for_display(time_list):
    """
    格式化时间列表用于显示
    """
    if not time_list:
        return ""

    return " | ".join([t.strftime("%H:%M") if isinstance(t, datetime) else str(t) for t in time_list])


def calculate_working_hours_with_details(time_list_normalized):
    """
    计算工作时间并返回详细信息
    """
    if not time_list_normalized or len(time_list_normalized) == 0:
        return {
            'total_hours': 0,
            'work_periods': [],
            'is_valid': False,
            'error': 'No valid times'
        }

    if len(time_list_normalized) % 2 != 0:
        return {
            'total_hours': 0,
            'work_periods': [],
            'is_valid': False,
            'error': f'Odd number of times: {len(time_list_normalized)}'
        }

    work_periods = []
    total_hours = 0

    for i in range(0, len(time_list_normalized), 2):
        start_time = time_list_normalized[i]
        end_time = time_list_normalized[i + 1]

        period_hours = (end_time - start_time).total_seconds() / 3600

        work_periods.append({
            'start': start_time.strftime("%H:%M"),
            'end': end_time.strftime("%H:%M"),
            'hours': round(period_hours, 2)
        })

        total_hours += period_hours

    return {
        'total_hours': round(total_hours, 2),
        'work_periods': work_periods,
        'is_valid': True,
        'error': None
    }