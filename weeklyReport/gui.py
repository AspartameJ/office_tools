from generateDoc import WeeklyReportGenerator
import sys
import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit

class WeeklyReportGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('周报生成器')
        self.setGeometry(300, 300, 600, 400)

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
        ]
        for label, widget in fields:
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            row.addWidget(widget)
            layout.addLayout(row)

        # 生成周报按钮
        self.generateButton = QPushButton('生成周报')
        self.generateButton.clicked.connect(self.generateReport)
        layout.addWidget(self.generateButton)

    def generateReport(self):
        # 这里应该调用你的WeeklyReportGenerator的逻辑
        # 示例中仅打印输入值
        year = self.yearEdit.text()
        week = self.weekEdit.text()
        month = self.monthEdit.text()
        day_start = self.dayStartEdit.text()
        day_end = self.dayEndEdit.text()
        name = self.nameEdit.text()
        project_names = self.projectNamesEdit.text().split('；')
        jobs_list = [job.split('；') for job in self.jobsListEdit.toPlainText().split('\n')]
        questions_list = self.questionsListEdit.toPlainText().split('；')
        plans_list = self.plansListEdit.toPlainText().split('；')

        print(year, week, month, day_start, day_end, name, project_names, jobs_list, questions_list, plans_list)
        # 这里添加调用WeeklyReportGenerator的代码
        report_generator = WeeklyReportGenerator(year, week, month, day_start, day_end, name)
        report_generator.generate_report(
            project_names=project_names,
            jobs_list=jobs_list,
            questions_list=questions_list,
            plans_list=plans_list
        )        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WeeklyReportGUI()
    ex.show()
    sys.exit(app.exec_())