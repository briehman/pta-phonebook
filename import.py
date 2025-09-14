import argparse
import openpyxl
import re
import sys
from itertools import groupby

class Student:
    def __init__(self, name, grade, teacher, guardians=None):
        self.name = name.strip()
        self.title = re.sub(r"(.+),\s+(.+)", r"\2 \1", self.name.strip()).title()
        self.grade = grade
        self.teacher = teacher
        self.guardians = guardians

    def __repr__(self):
        return str(self)

    def __hash__(self):
        return hash((self.grade, self.name))

    def __eq__(self, other):
        a = (self.grade, self.name)
        b = (other.grade, other.name)
        # print(f"checking {a} == {b} = {a == b}")
        return a == b

    def __lt__(self, other):
        return (self.grade, self.name) < (other.grade, other.name)

    def __str__(self):
        return f"{self.name} {self.title} - Grade {self.grade} - Teacher {self.teacher.title} - Guardians {self.guardians}"

    @staticmethod
    def parse_from_pta_file(row):
        has_phone = len(row) >= 7
        has_address = len(row) == 10

        grade = Grade(row[1].value)

        guardian = Guardian(
                name=row[4].value,
                email=row[5].value,
                phone=row[6].value if has_phone else None,
                address=row[7].value if has_address else None)

        teacher = Teacher(name=row[2].value, grade=grade)

        return Student(
                name=row[0].value,
                grade=grade,
                teacher=teacher,
                guardians=[guardian])

class Grade:
    def __init__(self, value):
        stripped = re.sub(r"(ST|ND|RD|TH|DG)", "", str(value)).strip()
        if stripped == 'K':
            self.grade = 'K'
            self.order = 0
        else:
            self.grade = int(float(stripped))
            self.order = self.grade

    def __lt__(self, other):
        return self.order < other.order

    def __hash__(self):
        return hash(self.order)

    def __repr__(self):
        return str(self)

    def __str__(self):
        return str(self.grade)

    def __eq__(self, other):
        return self.order == other.order

class Teacher:
    def __init__(self, name, grade):
        self.name = name
        if self.name == 'MORGAN EVANCIC':
            self.name = 'MORGAN BAETZ'
        self.grade = grade
        self.title = self.name.title()
        self.class_list_lookup = self.name.split(" ")[-1]
        self.students = []

    def __lt__(self, other):
        return (self.grade, self.name) < (other.grade, other.name)

    def __repr__(self):
        return str(self)

    def __str__(self):
        return f"{self.grade} - {self.title}"

    def __eq__(self, other):
        a = (self.grade, self.class_list_lookup)
        b = (other.grade, other.class_list_lookup)
        # print(f"checking {a} == {b} ? {a == b}")
        return a == b

    def add_students(self, students):
        self.students.extend(students)

class Class:
    def __init__(self, teacher, grade, students):
        self.teacher = teacher
        self.grade = grade
        self.students = students

    def __repr__(self):
        return str(self)

    def __str__(self):
        return f"{self.grade} - {self.teacher} - {self.students}"

class Guardian:
    def __init__(self, name, email, phone=None, address=None):
        self.name = name
        self.email = email.lower()
        self.phone = phone
        self.address = address.title() if address else None

    def title(self):
        return self.name.title()

    def __repr__(self):
        return str(self)

    def __str__(self):
        return f"{self.title()} {self.email} {self.phone} {self.address}"

class PtaParser:
    @staticmethod
    def parse_pta_students(pta_files):
        return [s for f in pta_files for s in PtaParser.__parse_pta_file(f)]

    @staticmethod
    def __parse_pta_file(f):
        wb = openpyxl.load_workbook(f)
        sheet = wb.active
        return [Student.parse_from_pta_file(row) for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column)]

class ClassListParser:
    @staticmethod
    def parse_class(sheet):
        teacher_name = re.sub(r" \(.*\)$", "", sheet.cell(1, 1).value.upper().replace('TEACHER: ', '').strip())
        # Transform Kdg, 1st, 2nd, 3rd, 4th, 5th => K, 1, 2, 3, 4, 5
        grade = Grade(sheet.cell(1, 2).value)

        if teacher_name != sheet.title:
            raise Exception(f"Expected teacher name {teacher_name} to match sheet title {sheet.title}")

        teacher = Teacher(teacher_name, grade)
        students = ClassListParser.parse_students(teacher, sheet)

        return Class(teacher, grade, students)

    @staticmethod
    def parse_students(teacher, sheet):
        students = []
        for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            if row[0] is None or row[0].value is None:
                break

            s = Student(name=row[0].value, grade=teacher.grade, teacher=teacher)
            students.append(s)
        return students

    @staticmethod
    def parse_lists(class_list):
        wb = openpyxl.load_workbook(class_list)

        teachers = []
        for sheet in wb.worksheets:
            if sheet.title.startswith("Sheet"):
                continue

            teacher = ClassListParser.parse_class(sheet)
            teachers.append(teacher)

        return teachers


parser = argparse.ArgumentParser(prog='PROG', usage='%(prog)s [options]')
parser.add_argument('--pta-files', nargs='+', help='the PTA directory files')
parser.add_argument('--class-list', help='the class list file')

args = parser.parse_args()

class_lists = ClassListParser.parse_lists(args.class_list)

students = PtaParser.parse_pta_students(args.pta_files)

for c in class_lists:
    pta_students = {s: s for s in students if s.teacher == c.teacher}

    print(c.teacher)
    class_students = [pta_students[s] if s in pta_students else s for s in c.students]

    print("\n".join(str(s) for s in class_students))
    print("")



# by_grade = lambda x: x.grade_order
# by_teacher = lambda x: x.teacher

# for grade, grade_students in groupby(sorted(students, key=by_grade), key=by_grade):
#     print(f"Grade: {grade}")

#     # print("\n".join(str(s) for s in sorted(gs, key=lambda x: x.teacher)))
#     for teacher, teacher_students in groupby(sorted(grade_students, key=by_teacher), key=by_teacher):
#         print(f"  Teacher: {teacher}")
#         print("\n".join(f"    {s}" for s in sorted(teacher_students, key=lambda x: x.name)))
#         print("")
#     print("")
