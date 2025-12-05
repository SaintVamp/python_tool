import datetime
import sys

import chinese_calendar
import pytz
from exchangelib import Credentials, Account, Configuration, DELEGATE, Message, HTMLBody, Mailbox
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter


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


account = sys.argv[1]
password = sys.argv[2]
beijing_tz = pytz.timezone('Asia/Shanghai')
today = datetime.datetime.now(beijing_tz).date()
print(f"今天的日期是: {today}")
result = get_next_workday(today)
if judge_workday(today):
    print('发送邮件')
    # 忽略SSL证书验证（因为可能是自签名证书）
    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

    # 使用NTLM验证的凭证
    credentials = Credentials(
        username=account,
        password=password  # 替换为实际密码
    )

    # 配置EWS服务
    config = Configuration(
        server='mail.newland.com.cn',  # 代理服务器地址
        credentials=credentials,
        auth_type='NTLM',  # 使用NTLM验证
    )

    # 连接到账户
    account = Account(
        primary_smtp_address=account,
        config=config,
        autodiscover=False,  # 已手动配置，不需要自动发现
        access_type=DELEGATE
    )
    with open('main_email', 'r', encoding='utf-8') as f:
        temps = f.readline().split(',')
    main_email = []
    for t in temps:
        main_email.append(Mailbox(email_address=t))
    with open('second_email', 'r', encoding='utf-8') as f:
        temps = f.readline().split(',')
    second_email = []
    for t in temps:
        second_email.append(Mailbox(email_address=t))
    message = Message(
        account=account,
        folder=account.sent,
        subject='【进门单】BOSS和网格通业务',
        body=HTMLBody(f'''
                    <html>
                        <body>
                            <p>您好，需要提进门单的人员如下：<br>
                                进门事由：项目日常沟通<br>
                                进门时间：{result}<br>
                                BOSS：张良、王鹏、朱翔宇、周小华、孙科、王香、朱春慧<br>
                                pulsar协查：曹鹏、王家豪、宋国栋、郭中奇<br>
                                网格通：孙健、李凌、戴波、周科</p>
                        </body>
                    </html>
                '''),
        to_recipients=main_email,
        cc_recipients=second_email
    )
    message.send()
