import pandas as pd
import numpy as np
from datetime import datetime
import holidays
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import traceback
import os
from utils.time_utils import get_minimum_distance, daily_working_time, parse_time_string, validate_time_format, normalize_time_list

class TimecardProcessor:
    def __init__(self, upload_folder, processed_folder):
        self.upload_folder = upload_folder
        self.processed_folder = processed_folder

    def process_step1(self, file_path):
        """Step1处理逻辑 - 包含完整的错误检测和高亮功能"""
        try:
            df = pd.read_excel(file_path)

            # 获取时间范围
            time_range = (df.iloc[1, 2]).replace("/", "").replace("~", "-").replace(" ", "")

            total_rows = df.shape[0]
            employee_amount = int((total_rows - 2) / 3)

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

            # 检测错误值位置
            error_value_location = []
            for i in range(employee_amount):
                for j in range(len(date_range)):
                    string = str(df_new.iloc[i, j + 1])

                    if string == 'nan':
                        continue
                    else:
                        letter = [x for x in string]
                        min_distance = get_minimum_distance(letter)
                        if min_distance == 3:
                            location = [i, j + 1]
                            error_value_location.append(location)

            # 检测奇数时间记录
            highlight_index = []
            for i in range(len(date_range)):
                time_rows = df_new[((df_new.iloc[:, i + 1].str.count(":") % 2) == 1).values].index.to_list()
                highlight_index.append(time_rows)

            # 合并错误位置
            for i in range(len(error_value_location)):
                r = error_value_location[i][0]
                c = error_value_location[i][1]
                highlight_index[c - 1].append(r)

            # 保存文件并添加高亮
            output_filename = f'table_with_error_cells({time_range}).xlsx'
            output_path = os.path.join(self.processed_folder, output_filename)
            df_new.to_excel(output_path, index=None, header=True)

            # 使用openpyxl添加高亮显示
            workbook = load_workbook(output_path)
            worksheet = workbook.active
            highlight_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

            # 高亮有问题的单元格
            for i in range(len(date_range)):
                for j in highlight_index[i]:
                    cell = worksheet.cell(row=j + 2, column=i + 2)
                    cell.fill = highlight_fill

            workbook.save(output_path)
            workbook.close()

            # 生成错误报告
            error_details = []
            total_highlighted = sum(len(rows) for rows in highlight_index)

            if error_value_location:
                error_details.append(f"发现 {len(error_value_location)} 个冒号距离异常的时间记录")

            if total_highlighted > len(error_value_location):
                odd_time_count = total_highlighted - len(error_value_location)
                error_details.append(f"发现 {odd_time_count} 个奇数时间记录")

            return {
                'success': True,
                'time_range': time_range,
                'output_file': output_filename,
                'employee_count': employee_amount,
                'error_count': len(error_value_location),
                'total_highlighted': total_highlighted,
                'error_details': error_details
            }

        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'traceback': traceback.format_exc()
            }

    def process_step2(self, error_file_path, time_range):
        """Step2处理逻辑 - 包含完整的工时计算、错误处理和多工作表生成"""
        try:
            df = pd.read_excel(error_file_path)
            df_new = df.copy()

            # 保存原始打卡时间数据
            df_original_times = df.copy()

            # 转换为字符串
            df = df.astype(str)

            # 记录有问题的数据
            problematic_data = []
            problematic_cells = []

            # 处理时间数据
            for i in range(df.shape[0]):
                for j in range(len(df.columns) - 1):
                    if (df.iloc[i, j + 1] == 'nan'):
                        continue

                    raw_time_str = str(df.iloc[i, j + 1])

                    # 分割时间字符串
                    time_list = parse_time_string(raw_time_str)

                    # 验证和规范化时间
                    time_list_normalized = normalize_time_list(time_list)
                    valid_times = len(time_list_normalized) > 0

                    # 检查是否有无效时间
                    for time_str in time_list:
                        if not validate_time_format(time_str):
                            problematic_data.append(
                                f"无效时间 - 员工: {df.iloc[i, 0]}, 列: {j + 1}, 时间: {time_str}")
                            problematic_cells.append((i, j + 1))
                            valid_times = False

                    # 计算工作时间
                    if valid_times and len(time_list_normalized) > 0:
                        if len(time_list_normalized) % 2 == 0:
                            df_new.iloc[i, j + 1] = daily_working_time(time_list_normalized)
                        else:
                            problematic_data.append(
                                f"奇数时间记录 - 员工: {df.iloc[i, 0]}, 列: {j + 1}, 时间数量: {len(time_list_normalized)}")
                            problematic_cells.append((i, j + 1))
                            df_new.iloc[i, j + 1] = 0
                    else:
                        df_new.iloc[i, j + 1] = 0

            # 计算工时统计
            n = len(df_new)

            # 第一周工时计算
            total1 = df_new.iloc[:, 1:8].sum(axis=1).to_list()
            HEG1 = [min(40, max(0, t)) if t > 0 else 0 for t in total1]
            OT1 = [max(0, t - 40) if t > 40 else 0 for t in total1]

            # 第二周工时计算
            total2 = df_new.iloc[:, 8:17].sum(axis=1).to_list()
            HEG2 = [min(40, max(0, t)) if t > 0 else 0 for t in total2]
            OT2 = [max(0, t - 40) if t > 40 else 0 for t in total2]

            # 总计
            Total_HEG = [HEG1[i] + HEG2[i] for i in range(n)]
            Total_OT = [OT1[i] + OT2[i] for i in range(n)]

            # 插入统计列
            df_new.insert(8, "HEG1", HEG1)
            df_new.insert(9, "OT1", OT1)
            df_new.insert(17, "HEG2", HEG2)
            df_new.insert(18, "OT2", OT2)
            df_new.insert(19, "Total_HEG", Total_HEG)
            df_new.insert(20, "Total_OT", Total_OT)

            # 识别需要检查迟到早退的员工
            name_list = df_new[(df_new['HEG1'] > 30) | (df_new['HEG2'] > 30)]['name'].to_list()

            # 处理原始时间数据用于显示
            df_original_for_display = df_original_times.astype(str).replace('nan', '')

            # 创建最终显示的数据框
            df_final = df_original_for_display.copy()
            original_date_cols = len(df_final.columns) - 1

            # 添加计算的工作小时数列
            for i in range(original_date_cols):
                col_name = f"{df_final.columns[i + 1]}_小时"
                df_final.insert(i + 1 + original_date_cols, col_name, df_new.iloc[:, i + 1])

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

            # 检测迟到早退
            attendance_result = self._detect_attendance_issues(df_original_for_display, name_list, df_final)

            # 处理假期
            holiday_result = self._process_holidays(time_range, df_final)

            # 创建Excel文件
            output_filename = f'work_attendance({time_range}).xlsx'
            output_path = os.path.join(self.processed_folder, output_filename)

            self._create_excel_report(df_final, df_original_for_display, attendance_result, 
                                    problematic_cells, original_date_cols, output_path, holiday_result)

            return {
                'success': True,
                'output_file': output_filename,
                'problematic_data': problematic_data,
                'problematic_cells_count': len(problematic_cells),
                'attendance_issues': attendance_result['attendance_issues'],
                'attendance_summary': attendance_result['attendance_summary'],
                'employee_count': len(df_final),
                'total_working_hours': sum(Total_HEG),
                'total_overtime': sum(Total_OT)
            }

        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'traceback': traceback.format_exc()
            }

    def _detect_attendance_issues(self, df_original_for_display, name_list, df_final):
        """检测迟到早退问题"""
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

                    # 提取时间字符串
                    time_snapshot = parse_time_string(raw_time_str)

                    # 验证时间格式
                    valid_times = []
                    for t in time_snapshot:
                        t = t.strip()
                        if t and validate_time_format(t):
                            valid_times.append(t)

                    time_snapshot = valid_times

                    if len(time_snapshot) == 0:
                        continue

                    try:
                        check_in_time = datetime.strptime(time_snapshot[0], '%H:%M')
                        morning_reference_time = datetime.strptime('10:00', '%H:%M')
                        check_out_time = datetime.strptime(time_snapshot[-1], '%H:%M')
                        evening_reference_time = datetime.strptime('17:00', '%H:%M')
                        check_times = len(time_snapshot)

                        employee_name = df_original_for_display.iloc[j, 0]
                        date_col = df_final.columns[i + 1]

                        # 检测迟到
                        if check_in_time > morning_reference_time:
                            highlight_rows_m.append(j)
                            attendance_issues.append(
                                f"迟到 - {employee_name}, {date_col}, 上班时间: {time_snapshot[0]}")

                        # 检测中午不打卡
                        if check_times == 2:
                            highlight_rows_n.append(j)
                            attendance_issues.append(
                                f"中午不打卡 - {employee_name}, {date_col}, 打卡次数: {check_times}")

                        # 检测早退
                        if check_out_time < evening_reference_time:
                            highlight_rows_e.append(j)
                            attendance_issues.append(
                                f"早退 - {employee_name}, {date_col}, 下班时间: {time_snapshot[-1]}")

                    except ValueError as e:
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
        holiday_mapping = {}  # 记录假期列名映射关系
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

        # 更新工时列名以匹配假期列名变化
        for old_col_name, new_col_name in holiday_mapping.items():
            old_hours_col = f"{old_col_name}_小时"
            new_hours_col = f"{new_col_name}_小时"
            if old_hours_col in df_final.columns:
                df_final = df_final.rename(columns={old_hours_col: new_hours_col})

        return {'holiday_column': holiday_column}

    def _create_excel_report(self, df_final, df_original_for_display, attendance_result, 
                           problematic_cells, original_date_cols, output_path, holiday_result):
        """创建Excel报告"""
        workbook = Workbook()
        workbook.remove(workbook.active)

        # 创建工作表
        sheet_names = ["时间汇总", "迟到", "中午不打卡", "早退"]
        sheets_data = [df_final, df_original_for_display, df_original_for_display, df_original_for_display]

        for sheet_name, data in zip(sheet_names, sheets_data):
            ws = workbook.create_sheet(title=sheet_name)
            for r_idx, row in enumerate(dataframe_to_rows(data, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

        # 定义颜色
        yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        problem_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

        # 高亮时间汇总工作表
        sheet1 = workbook["时间汇总"]
        metrics_start = len(df_final.columns) - 6
        for i in range(metrics_start, len(df_final.columns)):
            sheet1.cell(row=1, column=i + 1).fill = yellow_fill

        sheet1.cell(row=1, column=1).fill = red_fill
        if holiday_result['holiday_column'] is not None:
            sheet1.cell(row=1, column=holiday_result['holiday_column'] + 1).fill = green_fill

        # 标红有问题的数据
        for row_idx, col_idx in problematic_cells:
            if col_idx <= original_date_cols:
                sheet1.cell(row=row_idx + 2, column=col_idx + 1).fill = problem_fill

        # 高亮其他工作表
        c = df_original_for_display.shape[1]
        for sheet_idx, (sheet_name, highlight_cols) in enumerate(
                zip(["迟到", "中午不打卡", "早退"], 
                    [attendance_result['highlight_cols_m'], 
                     attendance_result['highlight_cols_n'], 
                     attendance_result['highlight_cols_e']])):
            sheet = workbook[sheet_name]
            sheet.cell(row=1, column=1).fill = red_fill

            for i in range(c - 1):
                for j in highlight_cols[i]:
                    sheet.cell(row=j + 2, column=i + 2).fill = red_fill

            # 标红有问题的数据
            for row_idx, col_idx in problematic_cells:
                if col_idx <= c - 1:
                    sheet.cell(row=row_idx + 2, column=col_idx + 1).fill = problem_fill

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
                adjusted_width = (max_length + 2) * 1.2
                sheet.column_dimensions[column_letter].width = adjusted_width

        workbook.save(output_path)
        workbook.close() 