from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt
import datetime

class WeeklyReportGenerator:
    def __init__(self, year=datetime.datetime.now().year, 
                 week=datetime.date.today().isocalendar()[1], 
                 month=datetime.datetime.now().month, 
                 day_start=datetime.datetime.now().day-4, 
                 day_end=datetime.datetime.now().day, 
                 name='杨佳'):
        self.year = year
        self.week = week
        self.month = month
        self.day_start = day_start
        self.day_end = day_end
        self.name = name
        self.doc = Document()
        self.setup_document()

    def setup_document(self):
        self.doc_name = f'系统与测试部AI开发部工作周报-{self.year}W{self.week}-{self.name}'
        self.greeding_str = f'戴总，宋哥，各位同事好！\n        这是{self.name}本周周报，请查收。'
        self.report_str = f'报告时间：{self.year}年{self.month}月{self.day_start}日—{self.year}年{self.month}月{self.day_end}日\n报告人： {self.name}'

    def add_greeting(self):
        greeding = self.doc.add_paragraph()
        greeding_runs = greeding.add_run(self.greeding_str)
        greeding_runs.font.size = Pt(12)

    def add_heading(self):
        heading = self.doc.add_paragraph()
        heading_runs = heading.add_run(self.doc_name)
        heading_runs.font.color.rgb = RGBColor(0,0,0)
        heading_runs.font.bold = True
        heading_runs.font.size = Pt(18)
        heading.paragraph_format.space_before = Pt(30)
        self.doc.paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_report_info(self):
        report = self.doc.add_paragraph()
        report.add_run(self.report_str)
        report.runs[0].font.size = Pt(12)

    def add_main_work(self, project_names, jobs_list):
        main_work = self.doc.add_paragraph()
        main_work.add_run('一、本周主要工作内容')
        main_work.runs[0].font.size = Pt(12)
        main_work.runs[0].font.bold = True
        for i, project_name in enumerate(project_names):
            project_name_runs = self.doc.add_paragraph().add_run(project_name)
            project_name_runs.font.bold = True
            project_name_runs.font.color.rgb = RGBColor(0,0,255)
            jobs = self.doc.add_paragraph()
            for job in jobs_list[i]:
                jobs.add_run(job + '\n')

    def add_questions(self, questions_list):
        questions = self.doc.add_paragraph()
        questions.add_run('二、存在问题')
        questions.runs[0].font.size = Pt(12)
        questions.runs[0].font.bold = True
        qss = self.doc.add_paragraph()
        for qs in questions_list:
            qss.add_run(qs + '\n')

    def add_next_week_plan(self, plans_list):
        new_work = self.doc.add_paragraph()
        new_work.add_run('三、下周重点工作计划')
        new_work.runs[0].font.size = Pt(12)
        new_work.runs[0].font.bold = True
        works = self.doc.add_paragraph()
        for work in plans_list:
            works.add_run(work + '\n')

    def set_font_style(self):
        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run.font._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    def save_document(self, filename="test.docx"):
        self.doc.save(filename)

    def generate_report(self, project_names, jobs_list, questions_list, plans_list):
        self.add_greeting()
        self.add_heading()
        self.add_report_info()
        self.add_main_work(project_names, jobs_list)
        self.add_questions(questions_list)
        self.add_next_week_plan(plans_list)
        self.set_font_style()
        self.save_document(self.doc_name + '.docx')

# 使用示例
if __name__ == "__main__":
    report_generator = WeeklyReportGenerator(year=2024, week=17, month=4, day_start=22, day_end=26, name='杨佳')
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