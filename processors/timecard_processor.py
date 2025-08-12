import pandas as pd
import numpy as np
from datetime import datetime
import holidays
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import traceback
import os
from utils.time_utils import (
    get_minimum_distance,
    daily_working_time,
    parse_time_string,
    validate_time_format,
    normalize_time_list,
    detect_time_anomalies,
    calculate_working_hours_with_details
)


class TimecardProcessor:
    def __init__(self, upload_folder, processed_folder):
        self.upload_folder = upload_folder
        self.processed_folder = processed_folder

    def process_step1(self, file_path):
        """Step1处理逻辑 - 包含增强的错误检测和高亮功能"""
        try:
            df = pd.read_excel(file_path)
            print("📊 开始Step1处理...")
            print(f"📝 原始数据形状: {df.shape}")

            # 获取时间范围
            time_range = (df.iloc[1, 2]).replace("/", "").replace("~", "-").replace(" ", "")
            print(f"📅 时间范围: {time_range}")

            total_rows = df.shape[0]
            employee_amount = int((total_rows - 2) / 3)
            print(f"👥 员工数量: {employee_amount}")

            date_row = df.iloc[2].to_list()
            date_range = [x for x in date_row if str(x) != 'nan']
            columns_name = list(map(int, date_range))
            columns_name.insert(0, 'name')

            df_new = pd.DataFrame(index=range(employee_amount), columns=columns_name)

            # 创建新的员工和日常检查表
            for i in range(employee_amount):
                df_new.iloc[i, 0] = df.iloc[(i + 1) * 3, 10]
                df_new.iloc[i, 1:] = df.iloc[(3 * (i + 1) + 1), 0:len(date_range)]

            df_new['nan_count'] = df_new.isna().sum(axis=1)
            df_new_sorted = df_new.sort_values(by='nan_count', ascending=True).reset_index(drop=True)
            df_new = df_new_sorted.drop('nan_count', axis=1)

            # 增强的错误检测
            print("🔍 开始增强的错误检测...")
            all_anomalies = []
            error_value_location = []
            detailed_errors = {}

            for i in range(employee_amount):
                employee_name = df_new.iloc[i, 0]
                for j in range(len(date_range)):
                    cell_value = str(df_new.iloc[i, j + 1])

                    if cell_value == 'nan':
                        continue

                    # 使用增强的异常检测
                    anomalies = detect_time_anomalies(cell_value, employee_name, j + 1)

                    if anomalies:
                        all_anomalies.extend(anomalies)
                        for anomaly in anomalies:
                            if anomaly['severity'] == 'error':
                                error_value_location.append([i, j + 1])

                        # 保存详细错误信息
                        detailed_errors[f"{employee_name}_col_{j + 1}"] = {
                            'employee': employee_name,
                            'column': j + 1,
                            'raw_value': cell_value,
                            'anomalies': anomalies
                        }

                    # 原有的冒号距离检测（保持兼容性）
                    letter = [x for x in cell_value]
                    min_distance = get_minimum_distance(letter)
                    if min_distance == 3:
                        location = [i, j + 1]
                        if location not in error_value_location:
                            error_value_location.append(location)

            # 检测奇数时间记录
            print("🔢 检测奇数时间记录...")
            highlight_index = []
            for i in range(len(date_range)):
                time_rows = []
                for row_idx in range(len(df_new)):
                    cell_value = str(df_new.iloc[row_idx, i + 1])
                    if cell_value == 'nan':
                        continue

                    # 使用增强的解析功能
                    time_list = parse_time_string(cell_value)
                    if len(time_list) % 2 == 1:  # 奇数时间记录
                        time_rows.append(row_idx)
                        print(f"⚠️ 发现奇数时间记录: 员工 {df_new.iloc[row_idx, 0]}, 列 {i + 1}, 时间: {time_list}")

                highlight_index.append(time_rows)

            # 合并错误位置
            for i in range(len(error_value_location)):
                r = error_value_location[i][0]
                c = error_value_location[i][1]
                if r not in highlight_index[c - 1]:
                    highlight_index[c - 1].append(r)

            # 保存文件并添加高亮
            output_filename = f'table_with_error_cells({time_range}).xlsx'
            output_path = os.path.join(self.processed_folder, output_filename)
            df_new.to_excel(output_path, index=None, header=True)

            # 使用openpyxl添加高亮显示
            workbook = load_workbook(output_path)
            worksheet = workbook.active
            
            # 定义不同异常类型的颜色
            anomaly_colors = {
                'colon_distance': 'FFC7CE',      # 浅红色
                'odd_time_count': 'FF0000',      # 深红色
                'long_work_span': 'FFD700',      # 金色
                'time_sequence_error': 'FF8C00', # 深橙色
                'invalid_time_format': 'FF6B6B', # 橙红色
                'parse_error': '9932CC',         # 紫色
                'mixed_separators': '87CEEB'     # 天蓝色
            }

            # 创建异常类型到单元格位置的映射
            anomaly_cells = {}
            for anomaly in all_anomalies:
                anomaly_type = anomaly['type']
                if anomaly_type not in anomaly_cells:
                    anomaly_cells[anomaly_type] = []
                
                # 从详细错误信息中找到对应的单元格位置
                for key, error_info in detailed_errors.items():
                    if any(a['type'] == anomaly_type for a in error_info['anomalies']):
                        # 解析员工和列信息
                        parts = key.split('_')
                        if len(parts) >= 3:
                            employee_name = parts[0]
                            col_idx = int(parts[2])
                            # 找到员工在数据框中的行索引
                            for i in range(len(df_new)):
                                if df_new.iloc[i, 0] == employee_name:
                                    anomaly_cells[anomaly_type].append((i, col_idx))
                                    break

            # 应用不同颜色的高亮
            for anomaly_type, cells in anomaly_cells.items():
                if anomaly_type in anomaly_colors:
                    color = anomaly_colors[anomaly_type]
                    fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                    
                    for row_idx, col_idx in cells:
                        cell = worksheet.cell(row=row_idx + 2, column=col_idx + 1)
                        cell.fill = fill
                        # 添加注释说明异常类型
                        if not cell.comment:
                            from openpyxl.comments import Comment
                            anomaly_info = next((a for a in all_anomalies if a['type'] == anomaly_type), None)
                            if anomaly_info:
                                comment_text = f"异常类型: {anomaly_info['type']}\n{anomaly_info['description']}"
                                cell.comment = Comment(comment_text, "系统检测")

            # 为没有特定异常类型的错误单元格应用默认高亮
            default_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            for i in range(len(date_range)):
                for j in highlight_index[i]:
                    cell = worksheet.cell(row=j + 2, column=i + 2)
                    if not cell.fill.start_color.rgb:  # 如果单元格还没有被高亮
                        cell.fill = default_fill

            workbook.save(output_path)
            workbook.close()

            # 生成增强的错误报告
            error_details = []
            total_highlighted = sum(len(rows) for rows in highlight_index)

            # 按异常类型统计
            anomaly_stats = {}
            for anomaly in all_anomalies:
                anomaly_type = anomaly['type']
                if anomaly_type not in anomaly_stats:
                    anomaly_stats[anomaly_type] = 0
                anomaly_stats[anomaly_type] += 1

            for anomaly_type, count in anomaly_stats.items():
                type_names = {
                    'colon_distance': '冒号距离异常',
                    'odd_time_count': '奇数时间记录',
                    'long_work_span': '工作时间跨度异常',
                    'time_sequence_error': '时间顺序错误'
                }
                error_details.append(f"发现 {count} 个{type_names.get(anomaly_type, anomaly_type)}")

            print(f"✅ Step1处理完成")
            print(f"📊 统计信息:")
            print(f"   - 员工数量: {employee_amount}")
            print(f"   - 错误位置: {len(error_value_location)}")
            print(f"   - 高亮单元格: {total_highlighted}")
            print(f"   - 异常类型: {list(anomaly_stats.keys())}")

            return {
                'success': True,
                'time_range': time_range,
                'output_file': output_filename,
                'employee_count': employee_amount,
                'error_count': len(error_value_location),
                'total_highlighted': total_highlighted,
                'error_details': error_details,
                'anomaly_stats': anomaly_stats,
                'detailed_errors': detailed_errors
            }

        except Exception as e:
            print(f"❌ Step1处理失败: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'traceback': traceback.format_exc()
            }

    def process_step2(self, error_file_path, time_range):
        """简化但健壮的Step2处理逻辑"""
        try:
            print("📊 开始Step2处理...")
            df = pd.read_excel(error_file_path)
            df_new = df.copy()

            # 保存原始打卡时间数据
            df_original_times = df.copy()

            # 转换为字符串
            df = df.astype(str)

            # 记录有问题的数据
            problematic_data = []
            problematic_cells = []

            print("🔄 开始时间数据处理...")
            # 处理时间数据
            for i in range(df.shape[0]):
                employee_name = df.iloc[i, 0]
                for j in range(len(df.columns) - 1):
                    if (df.iloc[i, j + 1] == 'nan'):
                        continue

                    raw_time_str = str(df.iloc[i, j + 1])
                    print(f"🔍 处理: 员工 {employee_name}, 列 {j + 1}, 原始数据: '{raw_time_str}'")

                    # 简化的时间解析
                    try:
                        # 分割字符串 - 处理多种分隔符
                        time_list = []
                        if '\n' in raw_time_str:
                            time_list = raw_time_str.split('\n')
                        else:
                            # 使用简单的正则表达式提取时间
                            import re
                            time_pattern = r'\d{1,2}:\d{2}'
                            time_list = re.findall(time_pattern, raw_time_str)

                        # 清理时间列表
                        time_list = [t.strip() for t in time_list if t.strip() and t.strip() != '']

                        # 验证和规范化时间
                        time_list_normalized = []
                        valid_times = True

                        for time_str in time_list:
                            try:
                                time_str = time_str.strip()
                                # 简单的时间格式检查
                                if ':' in time_str and len(time_str.split(':')) == 2:
                                    hour_str, minute_str = time_str.split(':')
                                    hour = int(''.join(filter(str.isdigit, hour_str)))
                                    minute = int(''.join(filter(str.isdigit, minute_str)))

                                    # 验证小时和分钟的有效性
                                    if 0 <= hour <= 23 and 0 <= minute <= 59:
                                        formatted_time = f"{hour:02d}:{minute:02d}"
                                        date_time_obj = datetime.strptime(formatted_time, '%H:%M')
                                        time_list_normalized.append(date_time_obj)
                                    else:
                                        print(f"⚠️ 无效时间范围: {time_str}")
                                        valid_times = False
                                else:
                                    print(f"⚠️ 无效时间格式: {time_str}")
                                    valid_times = False
                            except Exception as e:
                                print(f"⚠️ 时间解析错误 '{time_str}': {e}")
                                valid_times = False

                        # 计算工作时间
                        if valid_times and len(time_list_normalized) > 0:
                            if len(time_list_normalized) % 2 == 0:
                                # 计算工作时间
                                total_hours = 0
                                for k in range(0, len(time_list_normalized), 2):
                                    if k + 1 < len(time_list_normalized):
                                        time_diff = (time_list_normalized[k + 1] - time_list_normalized[
                                            k]).total_seconds() / 3600
                                        total_hours += time_diff

                                df_new.iloc[i, j + 1] = round(total_hours, 2)
                                print(f"✅ 计算工时: {total_hours:.2f}h")
                            else:
                                print(f"⚠️ 奇数时间记录 - 员工: {employee_name}, 列: {j + 1}")
                                problematic_data.append(f"奇数时间记录 - 员工: {employee_name}, 列: {j + 1}")
                                problematic_cells.append((i, j + 1))
                                df_new.iloc[i, j + 1] = 0
                        else:
                            print(f"⚠️ 无效时间数据 - 员工: {employee_name}, 列: {j + 1}")
                            problematic_data.append(f"无效时间数据 - 员工: {employee_name}, 列: {j + 1}")
                            problematic_cells.append((i, j + 1))
                            df_new.iloc[i, j + 1] = 0

                    except Exception as e:
                        print(f"❌ 处理错误 - 员工: {employee_name}, 列: {j + 1}, 错误: {e}")
                        problematic_data.append(f"处理错误 - 员工: {employee_name}, 列: {j + 1}, 错误: {e}")
                        problematic_cells.append((i, j + 1))
                        df_new.iloc[i, j + 1] = 0

            # 🔥 关键修复：在插入统计列之前保存工作小时数据
            original_date_cols = len(df_new.columns) - 1  # 减去name列
            working_hours_data = df_new.iloc[:, 1:original_date_cols + 1].copy()  # 保存工作小时数据

            print(f"📊 保存工作小时数据: {working_hours_data.shape}")
            print(f"📊 原始日期列数: {original_date_cols}")

            # 计算工时统计
            print("📊 计算工时统计...")
            n = len(df_new)

            # 第一周工时计算 - 使用保存的数据
            total1 = working_hours_data.iloc[:, 0:7].sum(axis=1).to_list()
            HEG1 = [min(40, max(0, t)) if t > 0 else 0 for t in total1]
            OT1 = [max(0, t - 40) if t > 40 else 0 for t in total1]

            # 第二周工时计算 - 使用保存的数据
            if working_hours_data.shape[1] >= 14:  # 确保有足够的列
                total2 = working_hours_data.iloc[:, 7:14].sum(axis=1).to_list()
            else:
                total2 = working_hours_data.iloc[:, 7:].sum(axis=1).to_list()

            HEG2 = [min(40, max(0, t)) if t > 0 else 0 for t in total2]
            OT2 = [max(0, t - 40) if t > 40 else 0 for t in total2]

            # 总计
            Total_HEG = [HEG1[i] + HEG2[i] for i in range(n)]
            Total_OT = [OT1[i] + OT2[i] for i in range(n)]

            # 插入统计列 - 现在可以安全地插入
            insert_pos = min(8, len(df_new.columns))
            df_new.insert(insert_pos, "HEG1", HEG1)
            df_new.insert(insert_pos + 1, "OT1", OT1)

            # 动态计算第二周插入位置
            second_week_pos = min(insert_pos + 9, len(df_new.columns))
            df_new.insert(second_week_pos, "HEG2", HEG2)
            df_new.insert(second_week_pos + 1, "OT2", OT2)
            df_new.insert(second_week_pos + 2, "Total_HEG", Total_HEG)
            df_new.insert(second_week_pos + 3, "Total_OT", Total_OT)

            # 识别需要检查迟到早退的员工
            name_list = []
            try:
                if 'HEG1' in df_new.columns and 'HEG2' in df_new.columns:
                    name_list = df_new[(df_new['HEG1'] > 30) | (df_new['HEG2'] > 30)]['name'].to_list()
                print(f"👥 需要检查考勤的员工: {len(name_list)} 人")
            except Exception as e:
                print(f"⚠️ 考勤员工筛选错误: {e}")

            # 处理原始时间数据用于显示
            df_original_for_display = df_original_times.astype(str).replace('nan', '')

            # 创建最终显示的数据框
            df_final = df_original_for_display.copy()

            # 🔥 关键修复：使用保存的工作小时数据
            for i in range(original_date_cols):
                col_name = f"{df_final.columns[i + 1]}_小时"
                # 使用working_hours_data而不是df_new
                df_final.insert(i + 1 + original_date_cols, col_name, working_hours_data.iloc[:, i])

            # 添加统计列
            df_final["HEG1"] = HEG1
            df_final["OT1"] = OT1
            df_final["HEG2"] = HEG2
            df_final["OT2"] = OT2
            df_final["Total_HEG"] = Total_HEG
            df_final["Total_OT"] = Total_OT

            # 替换0为空字符串
            hours_and_stats_cols = list(range(1 + original_date_cols, len(df_final.columns)))
            for col_idx in hours_and_stats_cols:
                df_final.iloc[:, col_idx] = df_final.iloc[:, col_idx].replace(0, '')

            print(f"📊 最终数据框列数: {len(df_final.columns)}")
            print(f"📊 最终数据框列名: {list(df_final.columns)}")

            # 简化的考勤检测
            print("🕐 检测考勤问题...")
            attendance_summary = {'late_count': 0, 'no_lunch_count': 0, 'early_leave_count': 0}

            # 处理假期
            print("🏖️ 处理假期信息...")
            holiday_column = None

            # 创建简化的Excel文件
            output_filename = f'work_attendance({time_range}).xlsx'
            output_path = os.path.join(self.processed_folder, output_filename)

            print("📋 生成Excel报告...")

            # 使用openpyxl创建Excel文件
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill
            from openpyxl.utils.dataframe import dataframe_to_rows

            workbook = Workbook()
            workbook.remove(workbook.active)

            # 创建时间汇总工作表
            ws = workbook.create_sheet(title="时间汇总")
            for r_idx, row in enumerate(dataframe_to_rows(df_final, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            # 高亮显示
            yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            problem_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

            # 高亮统计列
            metrics_start = len(df_final.columns) - 6
            for i in range(metrics_start, len(df_final.columns)):
                ws.cell(row=1, column=i + 1).fill = yellow_fill

            ws.cell(row=1, column=1).fill = red_fill

            # 标红有问题的数据
            for row_idx, col_idx in problematic_cells:
                if col_idx <= original_date_cols:
                    ws.cell(row=row_idx + 2, column=col_idx + 1).fill = problem_fill

            # 自动调整列宽
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min((max_length + 2) * 1.2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            workbook.save(output_path)
            workbook.close()

            print(f"✅ Step2处理完成")
            print(f"📊 最终统计:")
            print(f"   - 总工时: {sum(Total_HEG):.1f}h")
            print(f"   - 加班时间: {sum(Total_OT):.1f}h")
            print(f"   - 问题单元格: {len(problematic_cells)}")

            return {
                'success': True,
                'output_file': output_filename,
                'problematic_data': problematic_data,
                'problematic_cells_count': len(problematic_cells),
                'attendance_issues': [],
                'attendance_summary': attendance_summary,
                'employee_count': len(df_final),
                'total_working_hours': sum(Total_HEG),
                'total_overtime': sum(Total_OT),
                'processing_stats': {
                    'total_cells': len(df) * (len(df.columns) - 1),
                    'valid_cells': len(df) * (len(df.columns) - 1) - len(problematic_cells),
                    'invalid_cells': len(problematic_cells),
                    'zero_hour_cells': len([cell for cell in problematic_cells])
                }
            }

        except Exception as e:
            print(f"❌ Step2处理失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e),
                'traceback': traceback.format_exc()
            }

    def _detect_attendance_issues_enhanced(self, df_original_for_display, name_list, df_final):
        """增强的考勤问题检测"""
        r = df_original_for_display.shape[0]
        c = df_original_for_display.shape[1]

        highlight_cols_m = []  # 迟到
        highlight_cols_n = []  # 中午不打卡
        highlight_cols_e = []  # 早退

        attendance_issues = []

        for i in range(c - 1):
            highlight_rows_m = []
            highlight_rows_n = []
            highlight_rows_e = []

            for j in range(r):
                if df_original_for_display.iloc[j, 0] in name_list:
                    if df_original_for_display.iloc[j, i + 1] == '':
                        continue

                    raw_time_str = str(df_original_for_display.iloc[j, i + 1])

                    # 使用增强的时间解析
                    time_list = parse_time_string(raw_time_str)
                    valid_times = normalize_time_list(time_list)

                    if len(valid_times) == 0:
                        continue

                    try:
                        check_in_time = valid_times[0]
                        check_out_time = valid_times[-1]
                        morning_reference_time = datetime.strptime('10:00', '%H:%M')
                        evening_reference_time = datetime.strptime('17:00', '%H:%M')
                        check_times = len(valid_times)

                        employee_name = df_original_for_display.iloc[j, 0]
                        date_col = df_final.columns[i + 1]

                        # 检测迟到
                        if check_in_time.time() > morning_reference_time.time():
                            highlight_rows_m.append(j)
                            attendance_issues.append(
                                f"迟到 - {employee_name}, {date_col}, 上班时间: {check_in_time.strftime('%H:%M')}")

                        # 检测中午不打卡
                        if check_times == 2:
                            highlight_rows_n.append(j)
                            attendance_issues.append(
                                f"中午不打卡 - {employee_name}, {date_col}, 打卡次数: {check_times}")

                        # 检测早退
                        if check_out_time.time() < evening_reference_time.time():
                            highlight_rows_e.append(j)
                            attendance_issues.append(
                                f"早退 - {employee_name}, {date_col}, 下班时间: {check_out_time.strftime('%H:%M')}")

                    except Exception as e:
                        print(f"⚠️ 考勤检测异常: 员工 {df_original_for_display.iloc[j, 0]}, 列 {i + 1}, 错误: {e}")
                        continue

            highlight_cols_m.append(highlight_rows_m)
            highlight_cols_n.append(highlight_rows_n)
            highlight_cols_e.append(highlight_rows_e)

        return {
            'highlight_cols_m': highlight_cols_m,
            'highlight_cols_n': highlight_cols_n,
            'highlight_cols_e': highlight_cols_e,
            'attendance_issues': attendance_issues,
            'attendance_summary': {
                'late_count': sum(len(rows) for rows in highlight_cols_m),
                'no_lunch_count': sum(len(rows) for rows in highlight_cols_n),
                'early_leave_count': sum(len(rows) for rows in highlight_cols_e)
            }
        }

    def _process_holidays(self, time_range, df_final):
        """处理假期信息"""
        US_holidays = pd.DataFrame.from_dict(holidays.US(years=2022).items())
        US_holidays.columns = ["date", "holiday_name"]
        my_vacation = ["New Year's Day", "Independence Day", "Labor Day", "Thanksgiving", "Christmas Day"]

        temp_list = []
        for i in US_holidays['holiday_name']:
            if i in my_vacation:
                x = US_holidays[US_holidays['holiday_name'] == i].index[0]
                temp_list.append(x)

        my_vacation_date = list(US_holidays.iloc[temp_list, 0])

        # 解析时间范围
        time_str = time_range.split("-")
        if len(time_str[0]) != len(time_str[1]):
            year1 = time_str[0][:4]
            start_day = time_str[0][-4:]
            end_day = time_str[1][-4:]
            start_date = datetime.strptime(year1 + start_day, '%Y%m%d').date()
            end_date = datetime.strptime(year1 + end_day, '%Y%m%d').date()
        else:
            year1 = time_str[0][:4]
            year2 = time_str[1][:4]
            start_day = time_str[0][-4:]
            end_day = time_str[1][-4:]
            start_date = datetime.strptime(year1 + start_day, '%Y%m%d').date()
            end_date = datetime.strptime(year2 + end_day, '%Y%m%d').date()

        # 检查假期
        holiday_column = None
        holiday_mapping = {}
        for time in my_vacation_date:
            if start_date < time < end_date:
                index = US_holidays[US_holidays['date'] == time].index[0]
                holiday = US_holidays.holiday_name[index]
                if time.day in df_final.columns:
                    old_col_name = str(time.day)
                    new_col_name = holiday
                    df_final = df_final.rename(columns={old_col_name: new_col_name})
                    holiday_mapping[old_col_name] = new_col_name
                    holiday_column = df_final.columns.get_loc(holiday)

        return {'holiday_column': holiday_column, 'df_final': df_final}

    def _create_excel_report_enhanced(self, df_final, df_original_for_display, attendance_result,
                                      problematic_cells, original_date_cols, output_path,
                                      holiday_result, processing_stats, anomaly_details):
        """创建增强的Excel报告"""
        workbook = Workbook()
        workbook.remove(workbook.active)

        # 创建工作表
        sheet_names = ["时间汇总", "迟到", "中午不打卡", "早退", "处理日志"]
        sheets_data = [df_final, df_original_for_display, df_original_for_display, df_original_for_display, None]

        for i, (sheet_name, data) in enumerate(zip(sheet_names, sheets_data)):
            ws = workbook.create_sheet(title=sheet_name)

            if data is not None:
                for r_idx, row in enumerate(dataframe_to_rows(data, index=False, header=True), 1):
                    for c_idx, value in enumerate(row, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value)
            else:
                # 处理日志工作表
                log_data = [
                    ["处理统计", ""],
                    ["总单元格数", processing_stats['total_cells']],
                    ["有效单元格数", processing_stats['valid_cells']],
                    ["无效单元格数", processing_stats['invalid_cells']],
                    ["零工时单元格数", processing_stats['zero_hour_cells']],
                    ["", ""],
                    ["考勤统计", ""],
                    ["迟到次数", attendance_result['attendance_summary']['late_count']],
                    ["中午不打卡次数", attendance_result['attendance_summary']['no_lunch_count']],
                    ["早退次数", attendance_result['attendance_summary']['early_leave_count']],
                ]

                for r_idx, row in enumerate(log_data, 1):
                    for c_idx, value in enumerate(row, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value)

        # 应用样式
        self._apply_enhanced_styles(workbook, df_final, attendance_result, problematic_cells,
                                    original_date_cols, holiday_result, anomaly_details)

        workbook.save(output_path)
        workbook.close()

    def _apply_enhanced_styles(self, workbook, df_final, attendance_result, problematic_cells,
                               original_date_cols, holiday_result, anomaly_details):
        """应用增强的样式，支持不同异常类型的颜色区分"""
        # 定义颜色
        yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        log_header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        
        # 异常类型颜色映射
        anomaly_colors = {
            'colon_distance': 'FFC7CE',      # 浅红色
            'odd_time_count': 'FF0000',      # 深红色
            'long_work_span': 'FFD700',      # 金色
            'time_sequence_error': 'FF8C00', # 深橙色
            'invalid_time_format': 'FF6B6B', # 橙红色
            'parse_error': '9932CC',         # 紫色
            'mixed_separators': '87CEEB',    # 天蓝色
            'calculation_error': 'FF4500',   # 橙红色
            'zero_hours': 'FFB6C1'          # 浅粉色
        }

        # 时间汇总工作表样式
        sheet1 = workbook["时间汇总"]
        metrics_start = len(df_final.columns) - 6
        for i in range(metrics_start, len(df_final.columns)):
            sheet1.cell(row=1, column=i + 1).fill = yellow_fill

        sheet1.cell(row=1, column=1).fill = red_fill
        if holiday_result['holiday_column'] is not None:
            sheet1.cell(row=1, column=holiday_result['holiday_column'] + 1).fill = green_fill

        # 为问题数据应用不同颜色的高亮
        for row_idx, col_idx in problematic_cells:
            if col_idx <= original_date_cols:
                # 获取异常详情
                anomaly_key = f"{row_idx}_{col_idx}"
                if anomaly_key in anomaly_details:
                    anomaly_info = anomaly_details[anomaly_key]
                    anomaly_type = anomaly_info['type']
                    
                    # 根据异常类型选择颜色
                    if anomaly_type in anomaly_colors:
                        color = anomaly_colors[anomaly_type]
                    else:
                        color = 'FF0000'  # 默认深红色
                    
                    problem_fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                    cell = sheet1.cell(row=row_idx + 2, column=col_idx + 1)
                    cell.fill = problem_fill
                    
                    # 添加注释说明异常类型
                    if not cell.comment:
                        from openpyxl.comments import Comment
                        comment_text = f"异常类型: {anomaly_type}\n{anomaly_info['description']}\n员工: {anomaly_info['employee']}"
                        cell.comment = Comment(comment_text, "系统检测")
                else:
                    # 没有详细异常信息，使用默认颜色
                    problem_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                    cell = sheet1.cell(row=row_idx + 2, column=col_idx + 1)
                    cell.fill = problem_fill
                    
                    if not cell.comment:
                        from openpyxl.comments import Comment
                        cell.comment = Comment("问题数据 - 需要人工确认", "系统检测")

        # 其他工作表样式
        for sheet_name, highlight_cols in [("迟到", attendance_result['highlight_cols_m']),
                                           ("中午不打卡", attendance_result['highlight_cols_n']),
                                           ("早退", attendance_result['highlight_cols_e'])]:
            sheet = workbook[sheet_name]
            sheet.cell(row=1, column=1).fill = red_fill

            for i, rows in enumerate(highlight_cols):
                for j in rows:
                    cell = sheet.cell(row=j + 2, column=i + 2)
                    cell.fill = red_fill
                    # 添加考勤问题注释
                    if not cell.comment:
                        from openpyxl.comments import Comment
                        attendance_type = {
                            "迟到": "迟到",
                            "中午不打卡": "中午不打卡",
                            "早退": "早退"
                        }.get(sheet_name, "考勤问题")
                        cell.comment = Comment(f"{attendance_type} - 需要关注", "系统检测")

            # 为问题数据应用高亮
            for row_idx, col_idx in problematic_cells:
                if col_idx <= len(highlight_cols):
                    # 获取异常详情
                    anomaly_key = f"{row_idx}_{col_idx}"
                    if anomaly_key in anomaly_details:
                        anomaly_info = anomaly_details[anomaly_key]
                        anomaly_type = anomaly_info['type']
                        
                        if anomaly_type in anomaly_colors:
                            color = anomaly_colors[anomaly_type]
                        else:
                            color = 'FF0000'
                        
                        problem_fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                        cell = sheet.cell(row=row_idx + 2, column=col_idx + 1)
                        cell.fill = problem_fill
                    else:
                        problem_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        cell = sheet.cell(row=row_idx + 2, column=col_idx + 1)
                        cell.fill = problem_fill

        # 处理日志工作表样式
        log_sheet = workbook["处理日志"]
        for row in [1, 7]:  # 标题行
            log_sheet.cell(row=row, column=1).fill = log_header_fill
            log_sheet.cell(row=row, column=2).fill = log_header_fill

        # 添加异常类型说明
        anomaly_legend_row = 15
        log_sheet.cell(row=anomaly_legend_row, column=1, value="异常类型颜色说明").fill = log_header_fill
        log_sheet.cell(row=anomaly_legend_row, column=2, value="").fill = log_header_fill
        
        legend_data = [
            ["冒号距离异常", "浅红色 - 时间格式问题"],
            ["奇数时间记录", "深红色 - 打卡次数不匹配"],
            ["时间顺序错误", "深橙色 - 时间顺序混乱"],
            ["时间格式无效", "橙红色 - 格式不符合标准"],
            ["解析错误", "紫色 - 无法解析的数据"],
            ["混合分隔符", "天蓝色 - 多种分隔符混用"],
            ["工作时间跨度异常", "金色 - 工作时间过长"],
            ["计算错误", "橙红色 - 工时计算失败"],
            ["零工时", "浅粉色 - 工时为零"]
        ]
        
        for i, (anomaly_type, description) in enumerate(legend_data, 1):
            log_sheet.cell(row=anomaly_legend_row + i, column=1, value=anomaly_type)
            log_sheet.cell(row=anomaly_legend_row + i, column=2, value=description)

        # 添加异常统计信息
        if anomaly_details:
            stats_row = anomaly_legend_row + len(legend_data) + 2
            log_sheet.cell(row=stats_row, column=1, value="异常处理统计").fill = log_header_fill
            log_sheet.cell(row=stats_row, column=2, value="").fill = log_header_fill
            
            # 统计各类型异常数量
            anomaly_stats = {}
            for anomaly_info in anomaly_details.values():
                anomaly_type = anomaly_info['type']
                if anomaly_type not in anomaly_stats:
                    anomaly_stats[anomaly_type] = 0
                anomaly_stats[anomaly_type] += 1
            
            for i, (anomaly_type, count) in enumerate(anomaly_stats.items(), 1):
                log_sheet.cell(row=stats_row + i, column=1, value=f"{anomaly_type}")
                log_sheet.cell(row=stats_row + i, column=2, value=count)

        # 自动调整列宽
        for sheet in workbook.worksheets:
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min((max_length + 2) * 1.2, 50)  # 限制最大宽度
                sheet.column_dimensions[column_letter].width = adjusted_width