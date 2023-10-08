# coding=utf-8
import datetime
from typing import List

import pandas


class Student:
    def __init__(self, institute: str = '', major: str = '', class_name: str = '', name: str = '', student_id: int = 0,
                 nation: str = '', enroll_time: str = '190001', graduation_time: str = '190001', educate_time: int = 0,
                 birth_date: str = '19000101', identity_id: str = '', political_state: str = '', total_credit: int = 0,
                 compulsory_credit: int = 0, public_credit: int = 0, professional_credit: int = 0) -> None:
        self.institute = institute
        self.major = major
        self.class_name = class_name
        self.name = name
        self.student_id = student_id
        self.nation = nation
        self.enroll_time = datetime.datetime.strptime(enroll_time, '%Y%m')
        self.graduation_time = datetime.datetime.strptime(graduation_time, '%Y%m')
        self.educate_time = educate_time
        self.birth_date = datetime.datetime.strptime(birth_date, '%Y%m%d')
        self.identity_id = identity_id
        self.political_state = political_state
        self.total_credit = total_credit
        self.compulsory_credit = compulsory_credit
        self.public_credit = public_credit
        self.professional_credit = professional_credit

        self.course_list: List[Course] = []

    def __repr__(self):
        return f'<Student {self.name} {self.student_id}>'

    def __str__(self):
        return f'<Student {self.name} {self.student_id}>'

    def __gt__(self, other):
        return self.student_id > other.student_id

    @classmethod
    def load_from_pandas(cls, data_sheet) -> 'Student':
        institute = data_sheet.iloc[0, 1]
        major = data_sheet.iloc[0, 5]
        class_name = data_sheet.iloc[0, 10]
        name = data_sheet.iloc[1, 1]
        student_id = data_sheet.iloc[1, 5]
        nation = data_sheet.iloc[1, 10]
        enroll_time = data_sheet.iloc[2, 1]
        graduation_time = data_sheet.iloc[2, 5]
        educate_time = data_sheet.iloc[2, 10]
        birth_date = data_sheet.iloc[3, 1]
        identity_id = data_sheet.iloc[3, 5]
        political_state = data_sheet.iloc[3, 10]
        total_credit = data_sheet.iloc[4, 1]
        compulsory_credit = data_sheet.iloc[4, 3]
        public_credit = data_sheet.iloc[4, 7]
        professional_credit = data_sheet.iloc[4, 10]
        course_list = Course.load_from_pandas(data_sheet[6:len(data_sheet) - 1])

        student = cls(institute, major, class_name, name, student_id, nation, enroll_time, graduation_time,
                      educate_time, birth_date, identity_id, political_state, total_credit, compulsory_credit,
                      public_credit, professional_credit)
        student.course_list = course_list
        return student


class Course:
    def __init__(self, term: str = '', name: str = '', types: str = '', times: str = '', credit: float = 0,
                 score: str = '') -> None:
        self.term = term
        self.name = name
        self.types = types
        self.times = times
        self.credit = credit
        self.score = score

    def __repr__(self):
        return f'<Course {self.name} {self.score}>'

    def __str__(self):
        return f'<Course {self.name} {self.score}>'

    @classmethod
    def load_from_pandas(cls, data_sheet) -> List['Course']:
        course_list = []
        for index_x in range(0, 12, 7):
            for index_y in range(0, len(data_sheet)):
                course_term = data_sheet.iloc[index_y, 0 + index_x]
                course_name = data_sheet.iloc[index_y, 1 + index_x + (1 if index_x else 0)]
                course_type = data_sheet.iloc[index_y, 3 + index_x + (2 if index_x else 0)]
                course_time = data_sheet.iloc[index_y, 4 + index_x + (2 if index_x else 0)]
                course_credit = data_sheet.iloc[index_y, 5 + index_x + (2 if index_x else 0)]
                course_score = data_sheet.iloc[index_y, 6 + index_x + (2 if index_x else 0)]
                d = [course_term, course_name, course_type, course_time, course_credit, course_score]
                # if any d is nan
                if any(map(lambda x: pandas.isna(x), d)):
                    break
                course = cls(*d)
                course_list.append(course)
        return course_list

    @property
    def is_public(self) -> bool:
        return self.types == '公选'

    @property
    def is_professional(self) -> bool:
        return self.name in professional_courses_name


def increasing_number(data: pandas.DataFrame) -> int:
    return max(data) + 1


def load_student_data(path: str) -> List[Student]:
    raw_data = pandas.read_excel(path)
    student_list = []
    page_size = 61
    for index in range(0, len(raw_data), page_size):
        student_data = raw_data[index:index + page_size - 1]
        student = Student.load_from_pandas(student_data)
        student_list.append(student)
    return student_list


def main() -> None:
    data_sheet = '计科2211学生成绩卡.xls'
    summary_sheet = '工作簿1.xlsx'
    data = load_student_data(data_sheet)
    summary_sheet = pandas.read_excel(summary_sheet, header=0)
    # raw_columns = summary_sheet.columns
    summary_sheet.columns = range(0, len(summary_sheet.columns))
    # summary_dict = {}

    for student in data:
        for course in student.course_list:
            if '*' in course.score:
                print(student, course)
            elif course.is_professional:
                if int(course.score) < 70:
                    print(student, course)
            elif course.score.isnumeric():
                if int(course.score) < 60:
                    print(student, course)

        failed_course = []
        need_relearn_course = []
        unprofessional_course = []
        for course in student.course_list:
            score = course.score
            if '*' in score:
                score = score.replace('*', '')

            if not course.score.isnumeric():
                continue
            else:
                score = int(score)
            if course.is_professional:
                if score < 70:
                    failed_course.append(course)
                    unprofessional_course.append(course)
                    need_relearn_course.append(course)
            else:
                if score < 60:
                    failed_course.append(course)
                    need_relearn_course.append(course)

        if len(failed_course) + len(need_relearn_course) + len(unprofessional_course):
            # add a new line to datasheet
            loc = increasing_number(summary_sheet[0])
            summary_sheet.loc[loc, 0] = loc
            summary_sheet.loc[loc, 1] = student.student_id
            summary_sheet.loc[loc, 2] = student.name
            summary_sheet.loc[loc, 3] = student.major
            summary_sheet.loc[loc, 4] = student.enroll_time.year
            summary_sheet.loc[loc, 5] = '\n'.join([f'《{i.name}》{i.credit}学分' for i in failed_course])
            summary_sheet.loc[loc, 6] = f'{sum([float(i.credit) for i in failed_course])}学分'
            summary_sheet.loc[loc, 7] = '\n'.join([f'{i.name} {i.score}' for i in unprofessional_course])
            summary_sheet.loc[loc, 8] = '\n'.join([i.name for i in need_relearn_course])
    # summary_sheet.columns = raw_columns
    summary_sheet.to_excel('工作簿1.xlsx', index=False)
    pass


professional_courses_name = ['离散数学', '程序设计基础', '面向对象程序设计(Java)', '数据结构', '计算机组成与体系结构',
                             '操作系统原理', '计算机网络', '数据库系统原理']

if __name__ == '__main__':
    main()
