import smtplib
import os
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.header import Header
from parse import getEmail
from parse import classsifyClass
import parse
import datetime


dir_path = os.getcwd()
file_path = os.path.join(dir_path, '未打卡名单.xls')
name_path = os.path.join(dir_path, '18级学生信息_邮箱.xlsx')
df = getEmail(file_path, name_path)

mail_host = "smtp.exmail.qq.com"  # 使用邮箱的发送邮件服务器
mail_sender = "s_lyu@smail.nju.edu.cn"  # 发送人邮箱
mail_license = "wwhhcc110OK"  # 邮箱授权码


def sendMail(mm, mail_host, mail_sender, mail_receiver, mail_license):
    # 创建SMTP对象，采用SSL协议
    stp = smtplib.SMTP_SSL(mail_host, 465)
    # 设置发件人邮箱的域名和端口，端口地址为465
    stp.connect(mail_host, 465)
    # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
    stp.set_debuglevel(1)
    # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
    stp.login(mail_sender, mail_license)
    # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
    stp.sendmail(mail_sender, mail_receiver, mm.as_string())
    print("邮件发送成功")
    # 关闭SMTP对象
    stp.quit()


def send_warning(df, mail_host, mail_sender, mail_license):
    for index, rows in df.iterrows():
        mail_receiver = rows['邮箱']
        mm = MIMEMultipart('related')
        subject_content = "WARNING：您今日尚未在南大APP上打卡"
        mm["From"] = "电子18级年级工作小组<" + mail_sender + ">"
        mm["To"] = mail_receiver
        mm["Subject"] = Header(subject_content, 'utf-8')

        body_content = "      " + rows['姓名'] + \
            "同学，您今日(" + rows['填报日期'] + ")尚未在南大APP上打卡，请尽快填报！\r\n"
        if (rows['是否两天未打卡'] == "是"):
            body_content += "      此外，检测到您最近两天都未进行打卡，这将会影响到您的返校和评奖评优。\r\n"
        body_content += "      疫情期间，请做好日常防护工作，减少出行，不前往中高风险地区，祝您生活愉快。如果您已经打卡，请忽略本次邮件。\r\n    "
        message_text = MIMEText(body_content, "plain", "utf-8")
        mm.attach(message_text)
        sendMail(mm, mail_host, mail_sender, mail_receiver, mail_license)


def send_notice(df, mail_host, mail_sender, mail_receiver, mail_license):
    mm = MIMEMultipart('mixed')
    time = datetime.datetime.today()
    time = time.strftime("%Y-%m-%d")
    subject_content = "NOTICE：" + time + "未打卡名单"
    mm["From"] = "电子18级年级工作小组<" + mail_sender + ">"
    mm["To"] = ','.join(mail_receiver)
    mm["Subject"] = Header(subject_content, 'utf-8')
    html_msg = parse.get_df_html(df)
    content_html = MIMEText(html_msg, "html", "utf-8")
    mm.attach(content_html)
    sendMail(mm, mail_host, mail_sender, mail_receiver, mail_license)


if __name__ == '__main__':
    mail_receiver_notice = ['peiyu@nju.edu.cn', '359885114@qq.com', '2218315386@qq.com', '181180049@smail.nju.edu.cn', '2733147505@qq.com',
                            '2585608619@qq.com', '1084329358@qq.com', '181180164@smail.nju.edu.cn', '552418625@qq.com',
                            '1079450738@qq.com', '2543816228@qq.com', '15852756236@163.com', '2584922112@qq.com', '1679385376@qq.com']
    mail_receiver_notice_test = ['359885114@qq.com', 's_lyu@smail.nju.edu.cn']
    send_notice(df, mail_host, mail_sender,
                 mail_receiver_notice, mail_license)
    send_warning(df, mail_host, mail_sender, mail_license)
