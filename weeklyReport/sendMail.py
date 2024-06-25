import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from generateDoc import WeeklyReportGenerator

class EmailSender:
    def __init__(self, from_email, from_password, smtp_server, smtp_port):
        self.from_email = from_email
        self.from_password = from_password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def send_email(self, subject, body, to_emails, cc_emails=None, docx_filename=None, excel_filename=None):
        # 创建邮件对象
        msg = MIMEMultipart()
        msg['From'] = self.from_email
        msg['To'] = ', '.join(to_emails)
        if cc_emails:
            msg['Cc'] = ', '.join(cc_emails)
        msg['Subject'] = subject

        # 添加邮件正文（HTML格式）
        msg.attach(MIMEText(body, 'html'))

        # 添加docx附件
        if docx_filename:
            with open(docx_filename, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(docx_filename))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(docx_filename)}"'
            msg.attach(part)

        # 添加excel附件
        if excel_filename:
            with open(excel_filename, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(excel_filename))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_filename)}"'
            msg.attach(part)

        # 发送邮件
        server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
        server.login(self.from_email, self.from_password)
        text = msg.as_string()
        server.sendmail(self.from_email, to_emails + (cc_emails if cc_emails else []), text)
        server.quit()

# 示例用法
if __name__ == "__main__":
    from_email = "yangjia0395@fiberhome.com"
    from_password = "d5Njp2S625i7Kmwc"
    smtp_server = 'smtp.fiberhome.com'
    smtp_port = 465

    email_sender = EmailSender(from_email, from_password, smtp_server, smtp_port)

    subject = "系统与测试部AI开发部工作周报"
    to_emails = ["yangjia0395@yctco.com.cn"]
    cc_emails = ["1054136235@qq.com"]
    docx_filename = "系统与测试部AI开发部工作周报-2024W26-杨佳.docx"
    excel_filename = "系统与测试部AI开发部工作周报-2024W26-杨佳.xlsx"
    
    # 生成周报内容
    report_generator = WeeklyReportGenerator(year=2024, week=26, month=6, day_start=22, day_end=26, name='杨佳')
    project_names = ['项目一名称', '项目二名称']
    jobs_list = [
        ['任务1详情', '任务2详情'],
        ['任务1详情', '任务2详情', '任务3详情']
    ]
    questions_list = ['问题1详情', '问题2详情']
    plans_list = ['下周计划1详情', '下周计划2详情']

    report_generator.generate_report(
        project_names=project_names,
        jobs_list=jobs_list,
        questions_list=questions_list,
        plans_list=plans_list
    )
    
    # 获取HTML格式的周报内容
    body = report_generator.generate_html_report(
        project_names=project_names,
        jobs_list=jobs_list,
        questions_list=questions_list,
        plans_list=plans_list
    )
    
    email_sender.send_email(subject, body, to_emails, cc_emails, docx_filename, excel_filename)