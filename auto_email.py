import datetime
import sys
import os

import pytz
from exchangelib import Credentials, Account, Configuration, DELEGATE, Message, HTMLBody, Mailbox, FileAttachment
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

import tools

account = sys.argv[1]
password = sys.argv[2]
token = sys.argv[3]
beijing_tz = pytz.timezone('Asia/Shanghai')
today = datetime.datetime.now(beijing_tz).date()
print(f"今天的日期是: {today}")
print(f"操作时间: {datetime.datetime.now(beijing_tz).strftime('%H:%M:%S')}")
result = tools.get_next_workday(today)
data_arr = tools.get_mail_context(token,today)
mail_content=data_arr[0].replace('result', result)
user_list=data_arr[1]
attachment_path = r"来访信息导入文件.xls"
tools.generate_excel_with_data(attachment_path, user_list)
if tools.judge_workday(today):
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
        main_email.append(Mailbox(email_address=t.strip()))
    with open('second_email', 'r', encoding='utf-8') as f:
        temps = f.readline().split(',')
    second_email = []
    for t in temps:
        second_email.append(Mailbox(email_address=t.strip()))
    message = Message(
        account=account,
        folder=account.sent,
        subject='【进门单】BOSS和网格通业务',
        body=HTMLBody(mail_content),
        to_recipients=main_email,
        cc_recipients=second_email
    )
    if os.path.isfile(attachment_path):
        with open(attachment_path, 'rb') as f:
            content = f.read()
        filename = os.path.basename(attachment_path)
        file_attachment = FileAttachment(name=filename, content=content)
        message.attach(file_attachment)
        print(f"已添加附件: {filename}")
    else:
        print(f"警告: 附件路径不存在 {attachment_path}")
    message.send()
