from datetime import datetime
import re
import logging


def get_minimum_distance(letter):
    """计算冒号之间的最小距离"""
    try:
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
    except Exception as e:
        print(f"⚠️ get_minimum_distance error: {e}")
        return None


def daily_working_time(time_list_normalized):
    """计算每日工作时间"""
    try:
        if not time_list_normalized or len(time_list_normalized) == 0:
            return 0

        time_part = []
        i = len(time_list_normalized)

        while (i >= 1):
            if i >= 2:
                time_difference = (time_list_normalized[i - 1] - time_list_normalized[i - 2]).total_seconds() / (
                            60 * 60)
                i = i - 2
                time_part.append(time_difference)
            else:
                break

        return round(sum(time_part), 2)
    except Exception as e:
        print(f"⚠️ daily_working_time error: {e}")
        return 0


def parse_time_string(raw_time_str):
    """
    安全的时间字符串解析，包含完整的错误处理
    """
    try:
        if not raw_time_str or str(raw_time_str).strip() in ['nan', '', 'None', 'NaN']:
            return []

        raw_time_str = str(raw_time_str).strip()
        time_list = []

        print(f"🔍 解析时间字符串: '{raw_time_str}'")

        # 方法1: 处理换行分隔
        if '\n' in raw_time_str:
            lines = raw_time_str.split('\n')
            print(f"📝 换行分割: {lines}")

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # 如果行内还有空格分隔的时间，进一步分割
                if ' ' in line:
                    # 使用宽松的正则表达式提取时间
                    try:
                        time_pattern = r'\d{1,2}:\d{2}'
                        line_times = re.findall(time_pattern, line)
                        time_list.extend(line_times)
                        print(f"📝 行内空格分割: '{line}' -> {line_times}")
                    except Exception as e:
                        print(f"⚠️ 正则表达式错误: {e}")
                        # 如果正则失败，尝试手动分割
                        parts = line.split()
                        for part in parts:
                            if ':' in part and len(part.split(':')) == 2:
                                time_list.append(part.strip())
                else:
                    # 单行单个时间
                    if ':' in line and len(line.split(':')) == 2:
                        time_list.append(line)
                        print(f"📝 单行时间: '{line}'")

        # 方法2: 处理其他分隔符（空格、制表符、逗号等）
        elif any(sep in raw_time_str for sep in [' ', '\t', ',', ';']):
            try:
                # 先用正则表达式提取所有可能的时间
                time_pattern = r'\d{1,2}:\d{2}'
                time_matches = re.findall(time_pattern, raw_time_str)
                if time_matches:
                    time_list = time_matches
                    print(f"📝 正则提取（有分隔符）: {raw_time_str} -> {time_list}")
                else:
                    # 如果正则没匹配到，尝试分割
                    for separator in [' ', '\t', ',', ';', '  ', '   ']:
                        if separator in raw_time_str:
                            parts = raw_time_str.split(separator)
                            time_list = [p.strip() for p in parts if p.strip() and ':' in p]
                            break
                    print(f"📝 分隔符分割: {raw_time_str} -> {time_list}")
            except Exception as e:
                print(f"⚠️ 分隔符处理错误: {e}")
                # 手动处理
                parts = raw_time_str.replace('\t', ' ').replace(',', ' ').replace(';', ' ').split()
                time_list = [p.strip() for p in parts if p.strip() and ':' in p]

        # 方法3: 使用正则表达式提取所有时间格式（兜底方案）
        else:
            try:
                time_pattern = r'\d{1,2}:\d{2}'
                time_list = re.findall(time_pattern, raw_time_str)
                print(f"📝 正则提取（无分隔符）: {raw_time_str} -> {time_list}")
            except Exception as e:
                print(f"⚠️ 正则表达式错误: {e}")
                # 最后的手动尝试
                if ':' in raw_time_str:
                    time_list = [raw_time_str.strip()]

        # 清理和验证时间列表
        cleaned_times = []
        for time_str in time_list:
            try:
                time_str = str(time_str).strip()
                if time_str and time_str != '' and ':' in time_str:
                    # 基础格式验证 - 避免使用复杂的正则
                    parts = time_str.split(':')
                    if len(parts) == 2:
                        hour_str, minute_str = parts
                        # 移除非数字字符
                        hour_str = ''.join(filter(str.isdigit, hour_str))
                        minute_str = ''.join(filter(str.isdigit, minute_str))

                        if hour_str and minute_str:
                            hour = int(hour_str)
                            minute = int(minute_str)
                            if 0 <= hour <= 23 and 0 <= minute <= 59:
                                cleaned_time = f"{hour:02d}:{minute:02d}"
                                cleaned_times.append(cleaned_time)
                                print(f"✅ 有效时间: '{time_str}' -> '{cleaned_time}'")
                            else:
                                print(f"⚠️ 时间超出范围: '{time_str}' (hour={hour}, minute={minute})")
                        else:
                            print(f"⚠️ 无法提取数字: '{time_str}'")
                    else:
                        print(f"⚠️ 冒号分割失败: '{time_str}' -> {parts}")
                else:
                    print(f"⚠️ 跳过无效格式: '{time_str}'")
            except Exception as e:
                print(f"⚠️ 处理时间字符串时出错: '{time_str}' - {e}")
                continue

        print(f"🔄 最终清理后: {cleaned_times}")
        return cleaned_times

    except Exception as e:
        print(f"❌ parse_time_string 严重错误: {e}")
        return []


def validate_time_format(time_str):
    """安全的时间格式验证"""
    try:
        if not time_str or not isinstance(time_str, str):
            return False

        time_str = time_str.strip()

        # 避免使用复杂正则，直接检查格式
        if ':' not in time_str:
            return False

        parts = time_str.split(':')
        if len(parts) != 2:
            return False

        try:
            hour, minute = parts
            hour = int(hour)
            minute = int(minute)

            # 验证小时和分钟的有效性
            if 0 <= hour <= 23 and 0 <= minute <= 59:
                return True
        except ValueError:
            return False

        return False
    except Exception as e:
        print(f"⚠️ validate_time_format error: {e}")
        return False


def normalize_time_list(time_list):
    """
    安全的时间列表规范化
    """
    try:
        if not time_list:
            return []

        time_list_normalized = []
        invalid_times = []

        for time_str in time_list:
            try:
                time_str = str(time_str).strip()

                if not time_str or time_str in ['', 'nan', 'None']:
                    continue

                if validate_time_format(time_str):
                    try:
                        # 使用安全的datetime解析
                        date_time_obj = datetime.strptime(time_str, '%H:%M')
                        time_list_normalized.append(date_time_obj)
                        print(f"✅ 成功解析: {time_str}")
                    except ValueError as e:
                        invalid_times.append(f"{time_str} (解析错误: {e})")
                        print(f"⚠️ 时间解析失败: {time_str} - {e}")
                else:
                    invalid_times.append(f"{time_str} (格式无效)")
                    print(f"⚠️ 时间格式无效: {time_str}")
            except Exception as e:
                invalid_times.append(f"{time_str} (处理错误: {e})")
                print(f"⚠️ 时间处理错误: {time_str} - {e}")

        if invalid_times:
            print(f"🚨 发现 {len(invalid_times)} 个无效时间: {invalid_times}")

        print(f"✅ 成功解析 {len(time_list_normalized)} 个有效时间")
        return time_list_normalized

    except Exception as e:
        print(f"❌ normalize_time_list 严重错误: {e}")
        return []


def detect_time_anomalies(raw_time_str, employee_name, column_idx):
    """
    安全的时间数据异常检测
    """
    try:
        anomalies = []

        if not raw_time_str or str(raw_time_str).strip() in ['nan', '', 'None']:
            return anomalies

        raw_time_str = str(raw_time_str).strip()

        # 检测1: 冒号距离异常
        try:
            letter = [x for x in raw_time_str]
            min_distance = get_minimum_distance(letter)
            if min_distance == 3:
                anomalies.append({
                    'type': 'colon_distance',
                    'message': f'冒号距离异常 - 员工: {employee_name}, 列: {column_idx}, 最小距离: {min_distance}',
                    'severity': 'warning',
                    'color': 'FFC7CE',
                    'description': '时间格式问题，冒号前后数字位数异常'
                })
        except Exception as e:
            print(f"⚠️ 冒号距离检测错误: {e}")

        # 检测2: 解析时间
        try:
            time_list = parse_time_string(raw_time_str)
            valid_times = normalize_time_list(time_list)

            # 检测3: 奇数时间记录
            if len(valid_times) % 2 != 0:
                anomalies.append({
                    'type': 'odd_time_count',
                    'message': f'奇数时间记录 - 员工: {employee_name}, 列: {column_idx}, 时间数量: {len(valid_times)}',
                    'severity': 'error',
                    'color': 'FF0000',
                    'description': '打卡次数为奇数，无法配对计算工时'
                })

            # 检测4: 时间跨度异常
            if len(valid_times) >= 2:
                try:
                    time_span = (valid_times[-1] - valid_times[0]).total_seconds() / 3600
                    if time_span > 16:
                        anomalies.append({
                            'type': 'long_work_span',
                            'message': f'工作时间跨度异常 - 员工: {employee_name}, 列: {column_idx}, 跨度: {time_span:.1f}小时',
                            'severity': 'warning',
                            'color': 'FFD700',
                            'description': '单日工作时间跨度超过16小时，可能存在数据错误'
                        })
                except Exception as e:
                    print(f"⚠️ 时间跨度计算错误: {e}")

            # 检测5: 时间顺序异常
            if len(valid_times) >= 2:
                try:
                    for i in range(1, len(valid_times)):
                        if valid_times[i] <= valid_times[i - 1]:
                            anomalies.append({
                                'type': 'time_sequence_error',
                                'message': f'时间顺序异常 - 员工: {employee_name}, 列: {column_idx}, {valid_times[i - 1].strftime("%H:%M")} >= {valid_times[i].strftime("%H:%M")}',
                                'severity': 'error',
                                'color': 'FF8C00',
                                'description': '打卡时间顺序混乱，后一个时间早于前一个时间'
                            })
                            break
                except Exception as e:
                    print(f"⚠️ 时间顺序检测错误: {e}")

        except Exception as e:
            print(f"⚠️ 时间解析检测错误: {e}")
            anomalies.append({
                'type': 'parse_error',
                'message': f'解析错误 - 员工: {employee_name}, 列: {column_idx}, 原始数据: {raw_time_str}, 错误: {e}',
                'severity': 'error',
                'color': '9932CC',
                'description': f'时间字符串无法正确解析: {e}'
            })

        return anomalies

    except Exception as e:
        print(f"❌ detect_time_anomalies 严重错误: {e}")
        return []


def calculate_working_hours_with_details(time_list_normalized):
    """
    安全的工作时间计算
    """
    try:
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
            try:
                start_time = time_list_normalized[i]
                end_time = time_list_normalized[i + 1]

                period_hours = (end_time - start_time).total_seconds() / 3600

                work_periods.append({
                    'start': start_time.strftime("%H:%M"),
                    'end': end_time.strftime("%H:%M"),
                    'hours': round(period_hours, 2)
                })

                total_hours += period_hours
            except Exception as e:
                print(f"⚠️ 工时计算段错误: {e}")
                return {
                    'total_hours': 0,
                    'work_periods': [],
                    'is_valid': False,
                    'error': f'Period calculation error: {e}'
                }

        return {
            'total_hours': round(total_hours, 2),
            'work_periods': work_periods,
            'is_valid': True,
            'error': None
        }

    except Exception as e:
        print(f"❌ calculate_working_hours_with_details 严重错误: {e}")
        return {
            'total_hours': 0,
            'work_periods': [],
            'is_valid': False,
            'error': f'Calculation error: {e}'
        }