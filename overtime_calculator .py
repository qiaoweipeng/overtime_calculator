from calendar import month_name

import openpyxl
import re
import shutil
import os
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime


def calculate_overtime(input_excel_path, sheet_names=['深圳', '武汉', '长沙', '梅州'], output_dir='./output'):
    """
    计算Excel考勤表中指定工作表的加班时长，并生成核对报告。

    Args:
        input_excel_path (str): 输入Excel文件的路径。
        sheet_names (list): 需要处理的工作表名称列表，默认为['深圳', '武汉', '长沙', '梅州']。
        output_dir (str): 输出报告和带公式Excel文件的目录，默认为'./output'。
    """

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 复制原文件，不修改原始文件
    output_excel_path = os.path.join(output_dir,
                                     os.path.basename(input_excel_path).replace(".xlsx", "_加班核对版.xlsx"))
    shutil.copy(input_excel_path, output_excel_path)

    # 存储所有工作表的结果
    all_results = {}
    all_summary = {}

    # 先加载一次工作簿，处理所有工作表
    wb = openpyxl.load_workbook(output_excel_path)

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            print(f"警告：工作表 '{sheet_name}' 不存在，跳过处理。")
            continue

        ws_src = wb[sheet_name]

        # 读取列头
        headers = []
        for col in range(1, ws_src.max_column + 1):
            val = ws_src.cell(row=1, column=col).value
            headers.append(val)

        # 找出日期列（03-01 到 03-31）
        date_cols = []

        # --- 自动识别月份 START ---
        # 1. 先读取表头所有列名
        headers = []
        for col in range(1, ws_src.max_column + 1):
            val = ws_src.cell(row=1, column=col).value
            headers.append(val)

        # 2. 自动查找包含“月”的列名来确定月份
        target_month = None

        # 尝试多种匹配模式来识别月份 - 修复：支持单个数字月份 (5-01) 或双数字月份 (05-01)
        for header in headers:
            if header and isinstance(header, str):
                # 匹配类似 "5-01 星期五", "05-01 星期五", "5-01星期五", "05-01星期五" 的格式
                match = re.search(r'^(\d{1,2})-\d{2}', header.strip())  # 修复：使用 \d{1,2} 支持1位或2位月份
                if match:
                    month_num = match.group(1)
                    # 确保月份是两位数格式
                    target_month = f"{int(month_num):02d}"
                    print(f"✅ 成功自动识别到月份: {target_month}")
                    break

        # 3. 如果没识别到，抛出错误
        if not target_month:
            print(f"❌ 错误：无法在工作表 '{sheet_name}' 的表头中自动识别月份，请检查列名格式是否包含 'XX-XX'。")
            print(f"💡 提示：表头可能的格式包括 '03-01 星期日', '03-01星期日', '3-01 星期日' 等")
            print(f"📋 实际表头前几个元素: {headers[:10]}")  # 打印前10个表头元素用于调试
            continue
        # --- 自动识别月份 END ---

        # 4. 使用识别到的月份构建正则表达式（保留你之前修改的兼容空格逻辑）
        pattern = rf'{target_month}-\d{{2}}(\s?)星期[一二三四五六日]'

        # 查找所有匹配的日期列
        for col_idx, h in enumerate(headers):
            if h and re.match(pattern, str(h)):
                date_cols.append((col_idx + 1, h))  # 1-indexed

        # 如果没有找到匹配的日期列，尝试更宽松的匹配
        if not date_cols:
            print(f"⚠️  在工作表 '{sheet_name}' 中未找到匹配的日期列，尝试更宽松的匹配...")
            # 修复：也支持单数字月份格式的匹配
            loose_pattern = rf'{int(target_month)}-\d{{2}}|{target_month}-\d{{2}}'
            for col_idx, h in enumerate(headers):
                if h and re.match(loose_pattern, str(h)):
                    date_cols.append((col_idx + 1, h))

            if date_cols:
                print(f"✅ 找到 {len(date_cols)} 个可能的日期列")
            else:
                print(f"❌ 仍然无法找到日期列，跳过处理工作表 '{sheet_name}'")
                continue

        # 生成核对日期范围
        check_date_range = f"{target_month}月1日 — {target_month}月31日"

        # 找出'加班总时长/小时'列
        overtime_total_col = None
        for col_idx, h in enumerate(headers):
            if h and '加班总时长' in str(h):
                overtime_total_col = col_idx + 1
                break

        if not date_cols:
            print(f"错误：未找到日期列（如'03-01 星期日'）。请检查Excel文件格式。")
            continue
        if not overtime_total_col:
            print(f"警告：未找到'加班总时长/小时'列。将无法进行差异核对。")

        # ===== 定义辅助函数 =====
        def extract_last_time(cell_str):
            """
            从单元格字符串中提取最后一个有效的下班时间（HH:MM格式）。
            支持多种混合格式，如 "(HH:MM),(HH:MM)", "迟到X分钟(HH:MM),(HH:MM)", "缺卡(-),(HH:MM)"
            """
            if cell_str is None:
                return None

            cell_str = str(cell_str).strip()

            # 尝试匹配各种可能的下班时间模式
            # 优先匹配 (HH:MM) 格式
            all_times_in_parens = re.findall(r'\((\d{1,2}:\d{2})\)', cell_str)
            if all_times_in_parens:
                return all_times_in_parens[-1]

            # 如果没有 (HH:MM) 格式，但有类似 "缺卡(-),(20:12)" 这种，尝试提取
            match_missing_card = re.search(r'\(-\),(\d{1,2}:\d{2})', cell_str)
            if match_missing_card:
                return match_missing_card.group(1)

            return None

        def extract_overtime_start(cell_str):
            """从单元格备注提取 加班起算时间，格式：加班起算:21:00"""
            match = re.search(r'加班起算:(\d{1,2}:\d{2})', str(cell_str))
            if match:
                return match.group(1)
            return None

        def calc_overtime_minutes(end_time_str):
            """
            根据下班时间计算加班分钟数（按规则取整）。
            规则：19:00后才算加班，不足30分钟不算，超过30分钟按30分钟为单位向下取整。
            """
            if end_time_str is None:
                return 0

            try:
                h, m = map(int, end_time_str.split(':'))

                # 处理跨午夜情况（如00:07表示加班到第二天凌晨）
                # 假设所有打卡时间都在同一天，或凌晨时间属于前一天的加班
                if h < 7:  # 凌晨时间，认为是加班到第二天
                    end_minutes = 24 * 60 + h * 60 + m
                else:
                    end_minutes = h * 60 + m

                overtime_start = 19 * 60  # 19:00 = 1140 分钟

                if end_minutes <= overtime_start:
                    return 0

                raw_overtime = end_minutes - overtime_start

                # 按30分钟向下取整，不足30分钟不算
                units = raw_overtime // 30
                if units == 0:
                    return 0

                return units * 30

            except Exception:
                return 0

        # ===== 计算每个人的加班时长 =====
        results = []
        for row_idx in range(2, ws_src.max_row + 1):
            name = ws_src.cell(row=row_idx, column=1).value
            if name is None:
                continue

            dept = ws_src.cell(row=row_idx, column=2).value

            existing_total = None
            if overtime_total_col:
                existing_total = ws_src.cell(row=row_idx, column=overtime_total_col).value

            daily_details = []
            total_minutes = 0

            for col_idx, date_header in date_cols:
                cell_val = ws_src.cell(row=row_idx, column=col_idx).value
                if cell_val is None:
                    continue

                cell_str = str(cell_val)
                end_time = extract_last_time(cell_str)
                minutes = calc_overtime_minutes(end_time)
                total_minutes += minutes

                # 记录原始加班分钟数（未取整）用于详细报告
                raw_ot_min = 0
                if end_time:
                    h, m = map(int, end_time.split(':'))
                    end_min_val = (24 * 60 + h * 60 + m) if h < 7 else (h * 60 + m)
                    raw_ot_min = max(0, end_min_val - 19 * 60)

                daily_details.append({
                    'date': date_header,
                    'end_time': end_time,
                    'raw_overtime': raw_ot_min,
                    'minutes': minutes,
                    'raw_cell_data': cell_str
                })

            total_hours = total_minutes / 60

            # 比较与现有表格数据
            diff = None
            existing_float = None
            if existing_total is not None:
                try:
                    existing_float = float(existing_total)
                    diff = total_hours - existing_float
                except ValueError:
                    pass  # 无法转换为浮点数，保持diff为None

            results.append({
                'row_idx': row_idx,
                'name': name,
                'dept': dept,
                'total_minutes': total_minutes,
                'total_hours': total_hours,
                'existing_total': existing_total,
                'existing_float': existing_float,
                'diff': diff,
                'daily_details': daily_details
            })

        all_results[sheet_name] = results
        diff_items = [r for r in results if r['diff'] is not None and abs(r['diff']) > 0.001]
        all_summary[sheet_name] = {
            'total_people': len(results),
            'diff_count': len(diff_items),
            'check_date_range': f"{target_month}月1日 — {target_month}月31日"
        }

    # 样式定义
    header_fill_blue = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_fill_green = PatternFill(start_color='375623', end_color='375623', fill_type='solid')
    header_font_white = Font(color='FFFFFF', bold=True, size=10)
    diff_font_red = Font(color='CC0000', bold=True)
    diff_fill_light_red = PatternFill(start_color='FFE0E0', end_color='FFE0E0', fill_type='solid')
    ok_fill_light_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    alt_row_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=False)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # 为每个城市工作表生成核对列
    for sheet_name, results in all_results.items():
        ws = wb[sheet_name]

        # 添加新的列头
        last_col = ws.max_column
        ws.cell(row=1, column=last_col + 1, value="公式计算加班(小时)")
        ws.cell(row=1, column=last_col + 2, value="差异(小时)")
        ws.cell(row=1, column=last_col + 3, value="核对结果")

        # 为新列添加样式
        for col_idx in [last_col + 1, last_col + 2, last_col + 3]:
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill_blue
            cell.font = header_font_white
            cell.alignment = center_align
            cell.border = thin_border

        # 填充数据
        for i, r in enumerate(results):
            row_num = i + 2
            ws.cell(row=row_num, column=last_col + 1, value=r['total_hours'])
            if r['diff'] is not None:
                ws.cell(row=row_num, column=last_col + 2, value=r['diff'])
            else:
                ws.cell(row=row_num, column=last_col + 2, value="")

            is_diff = r['diff'] is not None and abs(r['diff']) > 0.001
            result_cell = ws.cell(row=row_num, column=last_col + 3, value='❌ 有差异' if is_diff else '✓ 一致')
            if is_diff:
                result_cell.fill = diff_fill_light_red
                result_cell.font = diff_font_red

    # ---- Sheet 1: 加班核对汇总 ----
    ws_summary = wb.create_sheet('加班核对汇总', 0)  # 插入到第一个位置
    ws_summary.title = '加班核对汇总'

    title_headers = ['城市', '序号', '姓名', '部门', '计算加班(小时)', '表格已有加班(小时)', '差异(小时)', '核对结果']
    col_widths = [10, 6, 10, 20, 16, 18, 12, 12]

    for col, (h, w) in enumerate(zip(title_headers, col_widths), 1):
        cell = ws_summary.cell(row=1, column=col, value=h)
        cell.fill = header_fill_blue
        cell.font = header_font_white
        cell.alignment = center_align
        cell.border = thin_border
        ws_summary.column_dimensions[get_column_letter(col)].width = w

    summary_row = 2
    for city, results in all_results.items():
        for i, r in enumerate(results):
            fill = alt_row_fill if (summary_row - 2) % 2 == 0 else PatternFill(fill_type=None)

            is_diff = r['diff'] is not None and abs(r['diff']) > 0.001

            values = [
                city,
                i + 1,
                r['name'],
                r['dept'],
                r['total_hours'],
                r['existing_float'] if r['existing_float'] is not None else r['existing_total'],
                f'{r["diff"]:.2f}' if r['diff'] is not None else 'N/A',
                '❌ 有差异' if is_diff else '✓ 一致'
            ]

            for col, val in enumerate(values, 1):
                cell = ws_summary.cell(row=summary_row, column=col, value=val)
                cell.border = thin_border
                cell.alignment = center_align if col not in [3, 4] else left_align

                if is_diff:
                    cell.fill = diff_fill_light_red
                    if col in [7, 8]:
                        cell.font = diff_font_red
                else:
                    cell.fill = fill

                if col in [5, 6, 7]:  # 小时数和差异格式化
                    cell.number_format = '0.00'
            summary_row += 1

    # ---- Sheet 2: 逐日加班明细 ----
    ws_detail = wb.create_sheet('逐日加班明细')

    detail_headers = ['城市', '姓名', '部门', '日期', '下班时间', '原始加班(分钟)', '计算加班(分钟)', '计算加班(小时)',
                      '原始数据']
    detail_widths = [10, 10, 20, 15, 10, 15, 15, 14, 50]

    for col, (h, w) in enumerate(zip(detail_headers, detail_widths), 1):
        cell = ws_detail.cell(row=1, column=col, value=h)
        cell.fill = header_fill_blue
        cell.font = header_font_white
        cell.alignment = center_align
        cell.border = thin_border
        ws_detail.column_dimensions[get_column_letter(col)].width = w

    detail_row_idx = 2
    for city, results in all_results.items():
        for r in results:
            for d in r['daily_details']:
                if d['minutes'] > 0:
                    values = [
                        city,
                        r['name'],
                        r['dept'],
                        d['date'],
                        d['end_time'],
                        d['raw_overtime'],
                        d['minutes'],
                        d['minutes'] / 60,
                        d['raw_cell_data']
                    ]
                    fill = alt_row_fill if detail_row_idx % 2 == 0 else PatternFill(fill_type=None)
                    for col, val in enumerate(values, 1):
                        cell = ws_detail.cell(row=detail_row_idx, column=col, value=val)
                        cell.border = thin_border
                        cell.alignment = center_align if col not in [2, 3, 9] else left_align
                        cell.fill = fill
                    ws_detail.cell(row=detail_row_idx, column=8).number_format = '0.00'
                    detail_row_idx += 1

    # ---- Sheet 3: 差异项目汇总 ----
    ws_diff = wb.create_sheet('差异项目汇总')

    diff_headers = ['城市', '序号', '姓名', '部门', '计算加班(小时)', '表格已有(小时)', '差异(小时)', '差异说明']
    diff_widths = [10, 6, 10, 20, 16, 16, 12, 50]

    for col, (h, w) in enumerate(zip(diff_headers, diff_widths), 1):
        cell = ws_diff.cell(row=1, column=col, value=h)
        cell.fill = header_fill_blue
        cell.font = header_font_white
        cell.alignment = center_align
        cell.border = thin_border
        ws_diff.column_dimensions[get_column_letter(col)].width = w

    diff_row_idx = 2
    has_any_diff = False
    for city, results in all_results.items():
        diff_items = [r for r in results if r['diff'] is not None and abs(r['diff']) > 0.001]
        if diff_items:
            has_any_diff = True
            for i, r in enumerate(diff_items):
                values = [
                    city,
                    i + 1,
                    r['name'],
                    r['dept'],
                    r['total_hours'],
                    r['existing_float'],
                    f'{r["diff"]:.2f}',
                    f'计算值{r["total_hours"]:.2f}h，表格值{r["existing_float"]:.2f}h，差异{r["diff"]:.2f}h'
                ]
                for col, val in enumerate(values, 1):
                    cell = ws_diff.cell(row=diff_row_idx, column=col, value=val)
                    cell.border = thin_border
                    cell.alignment = center_align if col not in [3, 4, 8] else left_align
                    cell.fill = diff_fill_light_red
                    if col in [5, 6, 7]:
                        cell.number_format = '0.00'
                diff_row_idx += 1

    if not has_any_diff:
        cell = ws_diff.cell(row=2, column=1, value='所有城市所有人员加班时长核对一致，无差异项目')
        cell.font = Font(color='008000', bold=True)
        ws_diff.merge_cells('A2:H2')
        cell.alignment = center_align

    # ---- Sheet 4: 加班核对说明 (原有的) ----
    # 找到并删除旧的说明工作表，确保更新
    if '加班核对说明' in wb.sheetnames:
        del wb['加班核对说明']
    if '单格公式说明' in wb.sheetnames:
        del wb['单格公式说明']

    ws_help = wb.create_sheet('加班核对说明')
    ws_help.column_dimensions['A'].width = 20
    ws_help.column_dimensions['B'].width = 70
    ws_help.column_dimensions['C'].width = 50

    # 标题
    title_cell = ws_help.cell(row=1, column=1, value='多城市考勤加班核对 — 使用说明')
    title_cell.font = Font(bold=True, size=14, color='1F4E79')
    ws_help.merge_cells('A1:C1')
    title_cell.alignment = Alignment(horizontal='center')

    # 说明内容
    help_data = [
        ('', '', ''),
        ('一、核对方法', '', ''),
        ('位置', '说明', ''),
        ('各城市工作表', '原始数据表，已在右侧新增3列核对列', ''),
        ('第54列', '公式计算加班(小时)：由程序按规则计算得出', ''),
        ('第55列', '差异(小时)：= 公式计算值 − 表格已有值', ''),
        ('第56列', '核对结果：✓ 一致 / ❌ 有差异', ''),
        ('', '', ''),
        ('二、加班计算规则', '', ''),
        ('规则', '说明', '示例'),
        ('起算时间', '19:00之后才算加班', '18:59下班 → 0分钟'),
        ('不足30分钟', '不计入加班', '19:25下班 → 0分钟（加了25分钟，不算）'),
        ('超30分钟不足1小时', '算30分钟', '19:55下班 → 30分钟（加了55分钟，算30分钟）'),
        ('超1小时不足1.5小时', '算1小时', '20:55下班 → 60分钟'),
        ('超1.5小时不足2小时', '算1.5小时', '21:25下班 → 90分钟'),
        ('跨午夜', '视为当天加班延续', '00:07下班 → 加班300分钟（5小时）'),
        ('', '', ''),
        ('三、为什么不能用普通Excel公式？', '', ''),
        ('原因', '说明', ''),
        ('数据格式为文本', '单元格内容如"(08:59),(23:04)"是文本字符串，不是时间值', ''),
        ('格式多样', '含迟到、缺卡、请假、外勤等多种混合格式', ''),
        ('需正则提取', '需要提取最后一个括号内的时间，普通公式难以实现', ''),
        ('', '', ''),
        ('四、如何快速定位差异？', '', ''),
        ('方法', '操作步骤', ''),
        ('筛选法', '在深圳工作表第56列（核对结果）开启筛选，选"❌ 有差异"即可', ''),
        ('条件格式', '差异行已标红色背景，可直接扫视定位', ''),
        ('', '', ''),
        ('五、单元格公式（可手动验证单个人）', '', ''),
        ('公式名称', '公式', '说明'),
        ('差异公式', '=公式计算列 - 加班总时长列', '结果为0表示一致'),
        ('核对公式', '=IF(ABS(差异)<0.01,"✓ 一致","❌ 有差异")', ''),
    ]

    row_num = 2
    for item in help_data:
        for col, val in enumerate(item, 1):
            if val:
                cell = ws_help.cell(row=row_num, column=col, value=val)
                if col == 1 and val.startswith(('一、', '二、', '三、', '四、', '五、')):
                    cell.font = Font(bold=True, color='1F4E79', size=11)
                elif row_num > 1 and item[0] in ('位置', '规则', '原因', '方法', '公式名称'):
                    cell.font = Font(bold=True)
        row_num += 1

    wb.save(output_excel_path)
    print(f'报告已保存至: {output_excel_path}')

    # 生成Markdown报告
    city_names = '_'.join(sheet_names)
    markdown_report_path = os.path.join(output_dir, f'{city_names}加班核对报告.md')
    with open(markdown_report_path, 'w', encoding='utf-8') as f:
        f.write("# 多城市考勤加班时长核对报告\n\n")
        f.write("**核对日期：** {}\n".format(datetime.now().strftime("%Y年%m月%d日")))
        f.write("**数据来源：** {}\n".format(os.path.basename(input_excel_path)))
        f.write("**核对城市：** {}\n\n---\n\n".format(', '.join(sheet_names)))

        total_people = sum([summary['total_people'] for summary in all_summary.values()])
        total_diffs = sum([summary['diff_count'] for summary in all_summary.values()])

        f.write("**统计汇总：**\n")
        f.write("- 总人数：{}人\n".format(total_people))
        f.write("- 差异项目：{}处\n\n".format(total_diffs))

        f.write("## 各城市核对情况\n\n")
        f.write("| 城市 | 核对人数 | 差异数量 | 核对范围 |\n")
        f.write("|------|---------|---------|----------|\n")
        for city, summary in all_summary.items():
            f.write(
                f"| {city} | {summary['total_people']}人 | {summary['diff_count']}处 | {summary['check_date_range']} |\n")
        f.write("\n---\n\n")

        f.write("## 核对规则说明\n\n")
        f.write("| 规则项目 | 说明 |\n")
        f.write("|---------|------|\n")
        f.write("| 加班起算时间 | 19:00之后才算加班 |\n")
        f.write("| 不足30分钟 | 不计入加班 |\n")
        f.write("| 超过30分钟不足1小时 | 计算为30分钟 |\n")
        f.write("| 超过1小时不足1.5小时 | 计算为1小时 |\n")
        f.write("| 计算单位 | 以30分钟为单位向下取整 |\n")
        f.write("| 跨午夜情况 | 凌晨时间（如00:07）视为当天加班延续 |\n\n---\n\n")

        f.write("## 核对结论\n\n")
        if total_diffs == 0:
            f.write("> **核对结果：全部{}人加班时长均与表格记录一致，无差异项目。**\n\n---\n\n".format(total_people))
        else:
            f.write("> **核对结果：发现 {} 处差异。请查看以下差异详情。**\n\n---\n\n".format(total_diffs))

        f.write("## 各城市加班时长汇总\n\n")
        for city, results in all_results.items():
            f.write(f"### {city}城市\n\n")
            f.write("| 序号 | 姓名 | 部门 | 计算加班(小时) | 表格已有(小时) | 差异(小时) | 核对结果 |\n")
            f.write("|-----|------|------|------------|------------|------------|--------|\n")
            for i, r in enumerate(results):
                is_diff = r['diff'] is not None and abs(r['diff']) > 0.001
                diff_str = f'{r["diff"]:.2f}' if r['diff'] is not None else 'N/A'
                existing_val = r['existing_float'] if r['existing_float'] is not None else r['existing_total']
                f.write(
                    f"| {i + 1} | {r['name']} | {r['dept']} | {r['total_hours']:.2f} | {existing_val} | {diff_str} | {'❌ 有差异' if is_diff else '✓ 一致'} |\n")
            f.write("\n")

        if total_diffs > 0:
            f.write("## 差异项目详情\n\n")
            f.write("| 城市 | 序号 | 姓名 | 部门 | 计算加班(小时) | 表格已有(小时) | 差异(小时) | 差异说明 |\n")
            f.write("|------|-----|------|------|------------|------------|------------|----------|\n")
            for city, results in all_results.items():
                diff_items = [r for r in results if r['diff'] is not None and abs(r['diff']) > 0.001]
                for i, r in enumerate(diff_items):
                    f.write(
                        f"| {city} | {i + 1} | {r['name']} | {r['dept']} | {r['total_hours']:.2f} | {r['existing_float']:.2f} | {r['diff']:.2f} | 计算值{r['total_hours']:.2f}h，表格值{r['existing_float']:.2f}h，差异{r['diff']:.2f}h |\n")
            f.write("\n---\n\n")

        # 加班时长较多人员（Top 10）
        all_results_flat = []
        for city, results in all_results.items():
            for r in results:
                r_with_city = r.copy()
                r_with_city['city'] = city
                all_results_flat.append(r_with_city)

        sorted_results = sorted(all_results_flat, key=lambda x: x['total_hours'], reverse=True)
        f.write("## 加班时长较多人员（Top 10）\n\n")
        f.write("| 排名 | 城市 | 姓名 | 部门 | 加班总时长(小时) |\n")
        f.write("|-----|------|------|------|-------------|\n")
        for i, r in enumerate(sorted_results[:10]):
            f.write(f"| {i + 1} | {r['city']} | {r['name']} | {r['dept']} | {r['total_hours']:.2f} |\n")
        f.write("\n---\n\n")

        f.write(
            "*报告由自动化脚本生成，核对逻辑：提取每日最后一次打卡时间，计算19:00后加班时长，以30分钟为单位向下取整（不足30分钟不计）。*\n")

    print(f'Markdown报告已保存至: {markdown_report_path}')


if __name__ == '__main__':
    # 示例用法
    # 请将 '3月考勤4.2_副本.xlsx' 放在与脚本相同的目录下，或者提供完整路径
    input_file = 'data.xlsx'
    # 默认处理深圳、武汉、长沙、梅州四个城市，如果不需要某个城市，从列表中删除即可
    cities_to_process = ['深圳', '武汉','长沙','梅州']
    output_folder = './output'

    calculate_overtime(input_file, sheet_names=cities_to_process, output_dir=output_folder)
