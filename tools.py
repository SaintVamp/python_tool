import datetime
import json

import chinese_calendar
import requests


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
    return body['results'][0]['properties']['context']['rich_text'][0]['plain_text']
