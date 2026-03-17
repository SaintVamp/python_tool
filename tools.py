import datetime
import json

import chinese_calendar
import pandas as pd
import requests
import xlwt


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
    mod_page(token, body['results'][0]['id'], body['results'][0]['properties']['date']['date']['start'])
    return [body['results'][0]['properties']['context']['rich_text'][0]['plain_text'], body['results'][0]['properties']['user']['rich_text'][0]['plain_text']]


def mod_page(token, page_id, ori_date):
    # API地址
    url = f"https://api.notion.com/v1/pages/{page_id}"

    # 请求头
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Notion-Version": "2025-09-03"  # 使用稳定的API版本
    }

    # 解析为日期对象（注意：原字符串只包含日期，没有时间）
    original_date = datetime.datetime.strptime(ori_date, "%Y-%m-%d").date()
    new_date = original_date + datetime.timedelta(days=30)
    new_date_str = new_date.isoformat()  # 格式: YYYY-MM-DD

    # 构造更新payload
    update_payload = {
        "properties": {
            "date": {  # 字段名必须与Notion中的一致（区分大小写）
                "date": {
                    "start": new_date_str,
                    "end": None,  # 如果不需要结束日期，保持null
                    "time_zone": None  # 时区也保持null
                }
            }
        }
    }

    # 发送PATCH请求
    requests.patch(url, headers=headers, json=update_payload)


def generate_excel_with_data(output_filename, visitor_data_text):
    columns = ['访客姓名', '来访单位', '身份证号', '联系电话']
    rows = visitor_data_text.split('；')
    visitor_list = []
    for row in rows:
        if row.strip():
            cells = [cell.strip() for cell in row.split('，')]
            if len(cells) >= 3:
                visitor_dict = {
                    '访客姓名': cells[0],
                    '来访单位': '新大陆',
                    '身份证号': cells[1],
                    '联系电话': cells[2]
                }
                visitor_list.append(visitor_dict)
    df = pd.DataFrame(visitor_list, columns=columns)
    if output_filename.lower().endswith('.xls'):
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('访客信息')
        for col, column_name in enumerate(columns):
            worksheet.write(0, col, column_name)
        for row, record in enumerate(visitor_list, start=1):
            for col, column_name in enumerate(columns):
                worksheet.write(row, col, record[column_name])
        workbook.save(output_filename)
        print(f"Excel文件已生成: {output_filename} (Excel 97-2003格式)")
    else:
        df.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"Excel文件已生成: {output_filename}")
    return output_filename
