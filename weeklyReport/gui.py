from generateDoc import WeeklyReportGenerator
from sendMail import EmailSender
import sys
import datetime
from PyQt5.QtWidgets import QFileDialog, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QProgressBar
import re


class WeeklyReportGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('周报生成器')
        self.setGeometry(300, 300, 1500, 400)

        layout = QVBoxLayout()
        self.setLayout(layout)

        # 获取当前日期和时间
        current_year = datetime.datetime.now().year
        current_week = datetime.date.today().isocalendar()[1]
        current_month = datetime.datetime.now().month
        current_day = datetime.datetime.now().day
        default_day_start = current_day - 4
        default_day_end = current_day
        default_name = '杨佳'

        # 输入字段，带默认值
        self.yearEdit = QLineEdit(str(current_year))
        self.weekEdit = QLineEdit(str(current_week))
        self.monthEdit = QLineEdit(str(current_month))
        self.dayStartEdit = QLineEdit(str(default_day_start))
        self.dayEndEdit = QLineEdit(str(default_day_end))
        self.nameEdit = QLineEdit(default_name)
        self.projectNamesEdit = QLineEdit()
        self.jobsListEdit = QTextEdit()
        self.questionsListEdit = QTextEdit()
        self.plansListEdit = QTextEdit()
        self.toEmailEdit = QLineEdit("weekreport@yctco.com.cn,songmeng@yctco.com.cn,daiquan@yctco.com.cn")
        self.ccEmailEdit = QLineEdit("yexin517@yctco.com.cn,dhxiao0394@yctco.com.cn,xuxiao@yctco.com.cn,tphe0469@yctco.com.cn,yswang0472@yctco.com.cn,hzhxiao9896@yctco.com.cn,yangjia0395@yctco.com.cn,jywang0027@yctco.com.cn,tlzhou0520@yctco.com.cn")

        # 添加到布局
        fields = [
            ('年份:', self.yearEdit),
            ('周数:', self.weekEdit),
            ('月份:', self.monthEdit),
            ('开始日期:', self.dayStartEdit),
            ('结束日期:', self.dayEndEdit),
            ('姓名:', self.nameEdit),
            ('项目名称（分号分隔）:', self.projectNamesEdit),
            ('任务详情（任务分号分隔，项目换行符分隔）:', self.jobsListEdit),
            ('存在问题（分号分隔）:', self.questionsListEdit),
            ('下周计划（分号分隔）:', self.plansListEdit),
            ('收件人（逗号分隔）:', self.toEmailEdit),
            ('抄送人（逗号分隔）:', self.ccEmailEdit),
        ]
        for label, widget in fields:
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            row.addWidget(widget)
            layout.addLayout(row)

        # 读取周报按钮
        self.readButton = QPushButton('1.读取周报')
        self.readButton.clicked.connect(self.readReport)
        layout.addWidget(self.readButton)

        # 生成周报按钮
        self.generateButton = QPushButton('2.生成周报')
        self.generateButton.clicked.connect(self.generateReport)
        layout.addWidget(self.generateButton)

        # 发送邮件按钮
        self.sendEmailButton = QPushButton('3.发送邮件')
        self.sendEmailButton.clicked.connect(self.sendEmail)
        layout.addWidget(self.sendEmailButton)

        # 进度条
        self.progressBar = QProgressBar(self)
        layout.addWidget(self.progressBar)

        # 报告生成器
        self.report_generator = WeeklyReportGenerator(self.yearEdit.text(), self.weekEdit.text(), self.monthEdit.text(), self.dayStartEdit.text(), self.dayEndEdit.text(), self.nameEdit.text())

    def generateReport(self):
        self.progressBar.setValue(0)
        self.progressBar.setFormat('生成周报中...')
        # 调用WeeklyReportGenerator的逻辑
        project_names = self.projectNamesEdit.text().split('；')
        jobs_list = [job.split('；') for job in self.jobsListEdit.toPlainText().split('\n')]
        questions_list = self.questionsListEdit.toPlainText().split('；')
        plans_list = self.plansListEdit.toPlainText().split('；')
        
        self.report_generator.generate_report(
            project_names=project_names,
            jobs_list=jobs_list,
            questions_list=questions_list,
            plans_list=plans_list
        )
        self.progressBar.setValue(100)
        self.progressBar.setFormat('生成周报完成')

    def sendEmail(self):
        self.progressBar.setValue(0)
        self.progressBar.setFormat('发送邮件中...')
        # 获取邮件内容
        project_names = self.projectNamesEdit.text().split('；')
        jobs_list = [job.split('；') for job in self.jobsListEdit.toPlainText().split('\n')]
        questions_list = self.questionsListEdit.toPlainText().split('；')
        plans_list = self.plansListEdit.toPlainText().split('；')

        # 生成HTML格式的周报内容
        html_body = self.report_generator.generate_html_report(
            project_names=project_names,
            jobs_list=jobs_list,
            questions_list=questions_list,
            plans_list=plans_list
        )
        self.progressBar.setValue(50)

        # 发送邮件
        from_email = "yangjia0395@fiberhome.com"
        from_password = "d5Njp2S625i7Kmwc"
        smtp_server = 'smtp.fiberhome.com'
        smtp_port = 465
        to_emails = self.toEmailEdit.text().split(',')
        cc_emails = self.ccEmailEdit.text().split(',')
        subject = self.report_generator.doc_name
        docx_filename = self.report_generator.doc_name + ".docx"
        excel_filename = self.report_generator.doc_name + ".xlsx"

        email_sender = EmailSender(from_email, from_password, smtp_server, smtp_port)
        email_sender.send_email(subject, html_body, to_emails, cc_emails, docx_filename, excel_filename)
        self.progressBar.setValue(100)
        self.progressBar.setFormat('发送邮件完成')

    def readReport(self):
        self.progressBar.setValue(0)
        self.progressBar.setFormat('读取周报中...')
        filename, _ = QFileDialog.getOpenFileName(self, '选择要读取的周报文件', '', 'All Files (*);;Text Files (*.txt)')
        if filename:
            report_data = WeeklyReportGenerator.read_report(filename)
            main_work = report_data['main_work']
            # 检查并设置项目名称
            project_names = '；'.join(list(main_work.keys()))
            self.projectNamesEdit.setText(project_names)

            # 检查并设置任务详情
            jobs_list = [(re.sub(r'\n+', '\n', value[0]).replace('\n','；')) for value in main_work.values()]
            jobs_list = [job[:-1] if job[-1] == '；' else job for job in jobs_list]
            jobs_str = '\n'.join(jobs_list)
            self.jobsListEdit.setPlainText(jobs_str)

            # 检查并设置存在问题
            questions_str = re.sub(r'\n+', '\n', report_data['questions'][0]).replace('\n','；')
            questions_str = questions_str[:-1] if questions_str[-1] == '；' else questions_str
            self.questionsListEdit.setPlainText(questions_str)

            # 检查并设置下周计划
            plans_str = re.sub(r'\n+', '\n', report_data['next_week_plan'][0]).replace('\n','；')
            plans_str = plans_str[:-1] if plans_str[-1] == '；' else plans_str
            self.plansListEdit.setPlainText(plans_str)
            self.progressBar.setValue(100)
            self.progressBar.setFormat('读取周报完成')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WeeklyReportGUI()
    ex.show()
    sys.exit(app.exec_())