import datetime
import json

import chinese_calendar
import requests
import pandas as pd


def judge_workday(day):
    if chinese_calendar.is_workday(day) and not chinese_calendar.is_holiday(day):
        return True
    return False


def get_next_workday(day):
    next_day = day + datetime.timedelta(days=1)
    while True:
        if chinese_calendar.is_workday(next_day) and not chinese_calendar.is_holiday(next_day):
            return f"{next_day.month}月{next_day.day}日"
        next_day += datetime.timedelta(days=1)


def get_mail_context(token, date_info):
    url = "https://api.notion.com/v1/data_sources/2da320ed-5837-804c-8a64-000bd00ed811/query"

    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2025-09-03",
        "Content-Type": "application/json"
    }

    data = {
        "filter": {
            "and": [
                {
                    "property": "date",
                    "date": {
                        "equals": f"{date_info}"
                    }
                }
            ]
        },
        "sorts": [
            {
                "property": "num_id",
                "direction": "ascending"
            }
        ]
    }
    response = requests.post(url, headers=headers, data=json.dumps(data))
    body = json.loads(response.text)
    return [body['results'][0]['properties']['context']['rich_text'][0]['plain_text'],body['results'][0]['properties']['user']['rich_text'][0]['plain_text']]


def generate_excel_with_data(output_filename, visitor_data_text):
    """
    生成包含访客信息的Excel文件
    :param output_filename: 输出的Excel文件名
    :param visitor_data_text: 包含访客数据的文本，用分号分割每行，用逗号分割每个单元格
    :return: 生成的Excel文件路径
    """
    # 定义列标题
    columns = ['访客姓名', '来访单位', '身份证号', '联系电话']
    
    # 解析文本数据
    rows = visitor_data_text.split('；')
    visitor_list = []
    
    for row in rows:
        if row.strip():  # 忽略空行
            cells = [cell.strip() for cell in row.split('，')]
            if len(cells) >= 3:
                visitor_dict = {
                    '访客姓名': cells[0],
                    '来访单位': '新大陆',
                    '身份证号': cells[1],
                    '联系电话': cells[2]
                }
                visitor_list.append(visitor_dict)
    
    # 创建DataFrame
    df = pd.DataFrame(visitor_list, columns=columns)
    
    # 保存为Excel文件
    df.to_excel(output_filename, index=False, engine='openpyxl')
    
    print(f"Excel文件已生成: {output_filename}")
    return output_filename