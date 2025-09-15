import argparse
import openpyxl
import re
import sys

from itertools import groupby
from openpyxl import Workbook
from openpyxl.styles import DEFAULT_FONT, Alignment, Border, Font, NamedStyle, Side


def fix_name(s):
    # Mcdonald -> McDonald
    return re.sub(r"Mc([a-z])", lambda m: "Mc" + m.group(1).upper(), s)

class Student:
    def __init__(self, name, grade, teacher, guardians=None):
        self.name = name.strip()
        self.title = fix_name(re.sub(r"(.+),\s+(.+)", r"\2 \1", self.name).title())
        self.index_name = fix_name(self.name.title())
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
        return a == b

    def __lt__(self, other):
        return (self.grade, self.name) < (other.grade, other.name)

    def __str__(self):
        return f"{self.name} {self.title} - Grade {self.grade} - Teacher {self.teacher.title} - Guardians {self.guardians}"

    def address(self):
        if self.guardians:
            addresses = [g.address for g in self.guardians if g.address]
            if addresses:
                return next((a for a in addresses if a is not None), None)
            else:
                return None

        else:
            return None

    @staticmethod
    def parse_from_pta_file(row, has_phone, has_address, guardian_2_index):
        grade = Grade(row[1].value)

        guardians = [Guardian(
                name=row[4].value,
                email=row[5].value,
                phone=row[6].value if has_phone else None,
                address=row[7].value if has_address else None)]

        if len(row) > guardian_2_index:
            if row[guardian_2_index].value:
                name2 = row[guardian_2_index].value if row[guardian_2_index] else None
                email2 = row[guardian_2_index + 1].value if row[guardian_2_index + 1] else None
                phone2 = row[guardian_2_index + 2].value if row[guardian_2_index + 2] else None
                guardians.append(Guardian(
                    name=name2,
                    email=email2,
                    phone=phone2,
                    ))

        teacher = Teacher(name=row[2].value, grade=grade)

        return Student(
                name=row[0].value,
                grade=grade,
                teacher=teacher,
                guardians=guardians)

class Grade:
    def __init__(self, value):
        stripped = re.sub(r"(ST|ND|RD|TH|DG)", "", str(value)).strip()
        if stripped == 'K':
            self.grade = 'K'
            self.order = 0
        else:
            self.grade = int(float(stripped))
            self.order = self.grade

    def pretty(self):
        match self.order:
            case 0: return "Kindergarten"
            case 1: return "1st Grade"
            case 2: return "2nd Grade"
            case 3: return "3rd Grade"
            case 4: return "4th Grade"
            case 5: return "5th Grade"
            case _: raise ValueError("Invalid grade")

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
        return a == b

    def add_students(self, students):
        self.students.extend(students)

class Class:
    def __init__(self, room, teacher, grade, students):
        self.room = re.sub(r"# ", "", room.title())
        self.teacher = teacher
        self.grade = grade
        self.students = students

    def title(self):
        return f"{self.grade.grade} {self.teacher.class_list_lookup}"

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
        if self.address:
            self.address = re.sub(r"Lombard, IL 60148", "", self.address, flags=re.IGNORECASE)

    def title(self):
        return fix_name(self.name.title())

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
        has_phone = sheet['G1'].value == 'Phone'
        has_address = sheet['H1'] and sheet['H1'].value == 'Address'
        if has_address:
            guardian_2_index = 10
        elif has_phone:
            guardian_2_index = 7
        else:
            guardian_2_index = 6

        return [Student.parse_from_pta_file(row, has_phone, has_address, guardian_2_index) for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column)]

class ClassListParser:
    @staticmethod
    def parse_class(sheet):
        teacher_name = re.sub(r" \(.*\)$", "", sheet.cell(1, 1).value.upper().replace('TEACHER: ', '').strip())
        # Transform Kdg, 1st, 2nd, 3rd, 4th, 5th => K, 1, 2, 3, 4, 5
        grade = Grade(sheet.cell(1, 2).value)
        room = sheet.cell(1, 3).value

        if teacher_name != sheet.title:
            raise Exception(f"Expected teacher name {teacher_name} to match sheet title {sheet.title}")

        teacher = Teacher(teacher_name, grade)
        students = ClassListParser.parse_students(teacher, sheet)

        return Class(room, teacher, grade, students)

    @staticmethod
    def parse_students(teacher, sheet):
        students = []
        for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            if row[0] is None or row[0].value is None or "total" in row[0].value.lower():
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


class TextOutput:

    blank_guardian = Guardian(name='', email='', phone='', address='')

    def print_class(self, cls):
        print(cls.teacher.title)
        print(f"{cls.grade.pretty()} - {cls.room}")

        for s in cls.students:
            address = s.address() if s.address() is not None else ''
            guardians = s.guardians
            if guardians:
                guardian1 = guardians[0]
                guardian2 = guardians[1] if len(guardians) > 1 else self.blank_guardian
            else:
                guardian1 = guardian2 = self.blank_guardian

            print(f"{s.title:30} {guardian1.title():30} {guardian1.email:30} {guardian1.phone if guardian1.phone else '':12}")
            if address:
                print(f"{"":4} {address:25} {guardian2.title():30} {guardian2.email:30} {str(guardian2.phone):12}")
        print("\n\n")


class ExcelOutput:

    blank_guardian = Guardian(name='', email='', phone='', address='')
    centered = Alignment(horizontal='center', vertical='center')


    def __init__(self):
        self.wb = openpyxl.Workbook()

        DEFAULT_FONT.name = 'Arial'
        DEFAULT_FONT.size = 10
        medium_border = Side(border_style='medium', color='000000')
        thin_border = Side(border_style='thin', color='000000')
        self.thin_border = thin_border

        heading = NamedStyle(name='heading')
        heading.alignment = Alignment(horizontal='center', vertical='center')
        heading.font = Font(name='Arial', bold=True, size=12)
        self.wb.add_named_style(heading)

        subheading = NamedStyle(name='subheading')
        subheading.alignment = Alignment(horizontal='center', vertical='center')
        subheading.font = Font(name='Arial', bold=True, size=11)
        self.wb.add_named_style(subheading)

        tableheading = NamedStyle(name='tableheading')
        tableheading.font = Font(name='Arial', bold=True, size=9)
        tableheading.border = Border(top=medium_border, bottom=medium_border)
        self.wb.add_named_style(tableheading)

        tableheadingend = NamedStyle(name='tableheadingend')
        tableheadingend.font = Font(name='Arial', bold=True, size=9)
        tableheadingend.border = Border(top=medium_border, bottom=medium_border, right=medium_border)
        self.wb.add_named_style(tableheadingend)

        student = NamedStyle(name='student')
        student.font = Font(name='Arial', bold=True, size=11)
        self.wb.add_named_style(student)

        studentend = NamedStyle(name='studentend')
        studentend.border = Border(right=thin_border)
        self.wb.add_named_style(studentend)

    def print_class(self, cls):
        print(f"Creating sheet {cls.title()}")

        ws = self.wb.create_sheet(title=cls.title())
        ws.merge_cells('A1:E1')
        ws.merge_cells('A2:E2')
        ws.column_dimensions['A'].width = 11
        ws.column_dimensions['B'].width = 22.5
        ws.column_dimensions['C'].width = 18.85
        ws.column_dimensions['D'].width = 31
        ws.column_dimensions['E'].width = 14

        ws['A1'] = cls.teacher.title
        ws['A1'].style = 'heading'
        ws['A2'] = f"{cls.grade.pretty()} - {cls.room}"
        ws['A2'].style = 'subheading'

        ws.append([])
        ws.append(['Student', 'Family Address', 'Parent/Guardian', 'Email', 'Phone'])
        ws['A4'].style = 'tableheading'
        ws['B4'].style = 'tableheading'
        ws['C4'].style = 'tableheading'
        ws['D4'].style = 'tableheading'
        ws['E4'].style = 'tableheadingend'

        idx = 5

        for s in cls.students:
            address = s.address() if s.address() is not None else ''
            guardians = s.guardians if s.guardians else []
            ws.insert_rows(idx=idx)
            ws.cell(row=idx, column=1)
            ws[f'A{idx}'] = s.title
            ws[f'A{idx}'].style = 'student'
            ws[f'E{idx}'].style = 'studentend'
            num_guardians = len(guardians)
            if num_guardians > 0:
                ws[f'C{idx}'] = guardians[0].title()
                ws[f'D{idx}'] = guardians[0].email
                ws[f'E{idx}'] = guardians[0].phone

                if num_guardians > 1 or address:
                    idx += 1
                    ws.insert_rows(idx=idx)
                    ws[f'E{idx}'].style = 'studentend'
                    if address:
                        ws[f'B{idx}'] = address
                    if num_guardians > 1:
                        ws[f'C{idx}'] = guardians[1].title()
                        ws[f'D{idx}'] = guardians[1].email
                        ws[f'E{idx}'] = guardians[1].phone
            # Put border on bottom
            ws[f'A{idx}'].border = Border(bottom=self.thin_border)
            ws[f'B{idx}'].border = Border(bottom=self.thin_border)
            ws[f'C{idx}'].border = Border(bottom=self.thin_border)
            ws[f'D{idx}'].border = Border(bottom=self.thin_border)
            ws[f'E{idx}'].border = Border(bottom=self.thin_border, right=self.thin_border)

            idx += 1

            # print(f"{s.title:30} {guardian1.title():30} {guardian1.email:30} {guardian1.phone if guardian1.phone else '':12}")
            # if address:
            #     print(f"{"":4} {address:25} {guardian2.title():30} {guardian2.email:30} {str(guardian2.phone):12}")


    def finish(self):
        self.wb.remove(self.wb.active)
        self.wb.save('output.xlsx')

        # for s in cls.students:
        #     address = s.address() if s.address() is not None else ''
        #     guardians = s.guardians
        #     if guardians:
        #         guardian1 = guardians[0]
        #         guardian2 = guardians[1] if len(guardians) > 1 else self.blank_guardian
        #     else:
        #         guardian1 = guardian2 = self.blank_guardian

        #     print(f"{s.title:30} {guardian1.title():30} {guardian1.email:30} {guardian1.phone if guardian1.phone else '':12}")
        #     if address:
        #         print(f"{"":4} {address:25} {guardian2.title():30} {guardian2.email:30} {str(guardian2.phone):12}")
        # print("\n\n")

parser = argparse.ArgumentParser(prog='PROG', usage='%(prog)s [options]')
parser.add_argument('--pta-files', nargs='+', help='the PTA directory files')
parser.add_argument('--class-list', help='the class list file')

args = parser.parse_args()

class_lists = ClassListParser.parse_lists(args.class_list)
students = PtaParser.parse_pta_students(args.pta_files)

for c in class_lists:
    pta_students = {s: s for s in students if s.teacher == c.teacher}

    class_students = sorted([pta_students[s] if s in pta_students else s for s in c.students])
    c.students = class_students

    class_students[0].teacher
    # The class list only lists the last name and room but the report includes the full name so use that version
    c.teacher = class_students[0].teacher

txt = TextOutput()
excel = ExcelOutput()
all_students = []
for c in class_lists:
    txt.print_class(c)
    excel.print_class(c)
    all_students.extend(c.students)

by_last_name_first_letter = lambda x: x.name[0]
students_by_letter = groupby(sorted(all_students, key=by_last_name_first_letter), key=by_last_name_first_letter)
for letter, students in students_by_letter:
    print(f"{letter}:")
    for s in sorted(list(students), key=lambda x: x.name):
        print(f"  {s.index_name:30} {s.grade} {s.teacher.class_list_lookup.title()}")
    print("")
print(len(all_students))

excel.finish()



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
