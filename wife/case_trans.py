import openpyxl
from openpyxl.utils import get_column_letter
import os
import re

# 字段映射：源字段 -> 目标字段（用于定位单元格）
FIELD_MAP = {
    '报案号': '报案号',
    '被保人': '被保险人',
    '报案日期': '报案日期',
    '追偿首次发起日期': '追偿发起日期',
    '诉讼届满日期': '诉讼时效届满日',
    '追偿案件类型': '追偿类型',
    '被追偿人名称': '被追偿人',
    '计划追偿金额': '计划追偿金额',
    '案情简介': '追偿过程简介'
}

# 目标单元格坐标（基于模板的固定位置）
TARGET_CELLS = {
    '报案号': 'C4',
    '被保险人': 'C2',
    '报案日期': 'E4',
    '追偿发起日期': 'E8',
    '诉讼时效届满日': 'E9',
    '追偿类型': 'C8',
    '被追偿人': 'C9',
    '计划追偿金额': 'C10',
    '追偿过程简介': 'C13'
}

def sanitize_sheet_name(name):
    """
    清洗工作表名称：移除非法字符，截断长度，避免重复由调用者处理。
    """
    if not name:
        name = "无名称"
    # 替换非法字符（: * ? / \ [ ] 等）
    name = re.sub(r'[\\/*?:\[\]]', '_', name)
    # 限制长度（openpyxl 最大 31）
    return name[:31]

def read_source_data(source_path):
    """
    读取数据源文件，返回行字典列表和列名映射。
    """
    wb = openpyxl.load_workbook(source_path, data_only=True)
    ws = wb.active

    # 获取标题行（第一行）
    headers = []
    for cell in ws[1]:
        headers.append(cell.value)

    # 建立列名到列索引的映射（1-based）
    col_index = {header: idx+1 for idx, header in enumerate(headers) if header}

    # 读取数据行（从第二行开始）
    data_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_dict = {}
        for header, idx in col_index.items():
            row_dict[header] = row[idx-1]  # values_only 返回的是0-based元组
        data_rows.append(row_dict)

    return data_rows, col_index

def fill_sheet(sheet, data_row):
    """
    根据数据行填充目标 sheet 的对应单元格。
    """
    for src_field, target_field in FIELD_MAP.items():
        if target_field in TARGET_CELLS:
            coord = TARGET_CELLS[target_field]
            value = data_row.get(src_field, '')
            # 如果值为 None，转为空字符串
            if value is None:
                value = ''
            sheet[coord] = value

def create_output_file(template_path, data_group, group_index):
    """
    根据一组数据生成一个输出文件（最多20条），
    group_index 用于文件编号（从1开始）。
    """
    # 加载模板工作簿
    template_wb = openpyxl.load_workbook(template_path)
    template_ws = template_wb.active

    # 创建新的输出工作簿
    output_wb = openpyxl.Workbook()
    # 删除默认生成的 Sheet
    output_wb.remove(output_wb.active)

    # 将模板工作表复制到新工作簿，作为空白模板（命名为 "_template"）
    template_copy = output_wb.create_sheet("_template")
    for row in template_ws.iter_rows():
        for cell in row:
            new_cell = template_copy.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection.copy()
                new_cell.alignment = cell.alignment.copy()

    # 用于记录已使用的 sheet 名称，避免重复
    used_names = set()

    for data_row in data_group:
        # 获取被追偿人名称作为 sheet 名
        raw_name = data_row.get('被追偿人名称', '')
        sheet_name = sanitize_sheet_name(raw_name)

        # 处理重名：添加后缀
        if sheet_name in used_names:
            suffix = 1
            while f"{sheet_name}_{suffix}" in used_names:
                suffix += 1
            sheet_name = f"{sheet_name}_{suffix}"
        used_names.add(sheet_name)

        # 复制模板工作表
        new_sheet = output_wb.copy_worksheet(output_wb["_template"])
        new_sheet.title = sheet_name

        # 填充数据
        fill_sheet(new_sheet, data_row)

    # 删除模板工作表
    output_wb.remove(output_wb["_template"])

    # 保存文件
    output_filename = f"目标表{group_index}.xlsx"
    output_wb.save(output_filename)
    print(f"已生成: {output_filename}")

def main():
    source_file = "数据源.xlsx"
    template_file = "目标表.xlsx"

    if not os.path.exists(source_file):
        print(f"错误：找不到数据源文件 {source_file}")
        return
    if not os.path.exists(template_file):
        print(f"错误：找不到模板文件 {template_file}")
        return

    # 读取所有数据行
    data_rows, _ = read_source_data(source_file)
    print(f"共读取 {len(data_rows)} 条数据")

    # 每20条一组
    group_size = 20
    groups = [data_rows[i:i+group_size] for i in range(0, len(data_rows), group_size)]

    for idx, group in enumerate(groups, start=1):
        create_output_file(template_file, group, idx)

    print("处理完成！")

if __name__ == "__main__":
    main()