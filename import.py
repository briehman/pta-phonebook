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

    def finish(self, data):
        for letter, students in data.students_index:
            print(f"{letter}:")
            for s in sorted(list(students), key=lambda x: x.name):
                print(f"  {s.index_name:30} {s.grade} {s.teacher.class_list_lookup.title()}")
            print("")

class ExcelOutput:

    blank_guardian = Guardian(name='', email='', phone='', address='')
    centered = Alignment(horizontal='center', vertical='center')


    def __init__(self):
        self.wb = openpyxl.Workbook()

        DEFAULT_FONT.name = 'Arial'
        DEFAULT_FONT.size = 10
        medium_border = Side(border_style='medium', color='000000')
        self.medium_border = medium_border
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

        indexletter = NamedStyle(name='indexletter')
        indexletter.alignment = Alignment(horizontal='center', vertical='bottom')
        indexletter.font = Font(name='Arial', bold=True, size=9)
        self.wb.add_named_style(indexletter)

        indexstudent = NamedStyle(name='indexstudent')
        indexstudent.font = Font(name='Arial', size=10)
        indexstudent.border = Border(top=thin_border, right=thin_border, bottom=thin_border, left=thin_border)
        self.wb.add_named_style(indexstudent)

        self.create_welcome()


    def create_welcome(self):
        welcome = self.wb.create_sheet(title='Welcome')
        welcome.column_dimensions['A'].width = 97

        welcome['A2'] = 'PTA PHONE BOOK'
        welcome['A2'].font = Font(name='Arial', bold=True, size=24)
        welcome['A2'].alignment = Alignment(horizontal='center')

        welcome['A3'] = '2025-2026'
        welcome['A3'].font = Font(name='Arial', bold=True, size=14)
        welcome['A3'].alignment = Alignment(horizontal='center')

        welcome['A6'] = 'WILLIAM HAMMERSCHMIDT SCHOOL'
        welcome['A6'].font = Font(name='Arial', bold=True, size=20)
        welcome['A6'].alignment = Alignment(horizontal='center')

        welcome['A7'] = '617 Hammerschmidt Avenue'
        welcome['A7'].font = Font(name='Arial', size=14)
        welcome['A7'].alignment = Alignment(horizontal='center')

        welcome['A8'] = 'Lombard, IL 60148'
        welcome['A8'].font = Font(name='Arial', size=14)
        welcome['A8'].alignment = Alignment(horizontal='center')

        welcome['A11'] = 'Phone: 630-827-4200    Fax: 630-620-3733'
        welcome['A11'].font = Font(name='Arial', size=14)
        welcome['A11'].alignment = Alignment(horizontal='center')

        welcome['A13'] = 'VOICEMAIL/ATTENDANCE 630-827-4201'
        welcome['A13'].font = Font(name='Arial', size=14)
        welcome['A13'].alignment = Alignment(horizontal='center')

        welcome['A16'] = 'School District 44'
        welcome['A16'].font = Font(name='Arial', size=14)
        welcome['A16'].alignment = Alignment(horizontal='center')

        welcome['A17'] = 'Website: www.sd44.org'
        welcome['A17'].font = Font(name='Arial', size=14)
        welcome['A17'].alignment = Alignment(horizontal='center')

        welcome['A19'] = 'Mr. David Danielski'
        welcome['A19'].font = Font(name='Arial', size=14)
        welcome['A19'].alignment = Alignment(horizontal='center')

        welcome['A20'] = 'PRINCIPAL'
        welcome['A20'].font = Font(name='Arial', size=14)
        welcome['A20'].alignment = Alignment(horizontal='center')

        welcome['A22'] = 'Ms. Liz Valdivia'
        welcome['A22'].font = Font(name='Arial', size=14)
        welcome['A22'].alignment = Alignment(horizontal='center')

        welcome['A23'] = 'SECRETARY'
        welcome['A23'].font = Font(name='Arial', size=14)
        welcome['A23'].alignment = Alignment(horizontal='center')

        welcome['A33'] = "THIS PTA PHONE BOOK IS FOR PARENT AND STUDENT USE ONLY,\nNOT FOR COMMERCIAL USE."
        welcome['A33'].font = Font(name='Arial', size=14)
        welcome['A33'].alignment = Alignment(horizontal='center', wrapText=True)

        welcome['A35'] = 'This Phone Book is sponsored by the WHS PTA and is issued free, one per member family.'
        welcome['A35'].font = Font(name='Arial', size=12)
        welcome['A35'].alignment = Alignment(horizontal='center')

    def print_class(self, cls):
        ws = self.wb.create_sheet(title=cls.title())
        ws.merge_cells('A1:E1')
        ws.merge_cells('A2:E2')
        ws.column_dimensions['A'].width = 11
        ws.column_dimensions['B'].width = 22.5
        ws.column_dimensions['C'].width = 19
        ws.column_dimensions['D'].width = 28
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
                ws[f'D{idx}'].alignment = Alignment(wrap_text=True)
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

    def finish(self, data):
        self.wb.remove(self.wb.active)

        ws = self.wb.create_sheet(title='Index')
        ws.merge_cells('A1:G1')
        ws.merge_cells('A2:G2')
        ws['A1'] = 'STUDENT INDEX'
        ws['A1'].style = 'heading'
        ws['A1'].font = Font(name='Arial', bold=True, size=11)
        ws['A1'].border = Border(bottom=self.medium_border)
        ws['A2'] = 'Last Name, First Name, Grade, Teacher'
        ws['A2'].style = 'subheading'
        ws['A2'].font = Font(name='Arial', bold=True, size=10)

        name_width = 27.8
        grade_width = 3.85
        teacher_width = 14.42

        ws.column_dimensions['A'].width = name_width
        ws.column_dimensions['B'].width = grade_width
        ws.column_dimensions['C'].width = teacher_width
        ws.column_dimensions['D'].width = 4.7
        ws.column_dimensions['E'].width = name_width
        ws.column_dimensions['F'].width = grade_width
        ws.column_dimensions['G'].width = teacher_width

        pos = ExcelIndexPositioner(columns=[('A', 'B', 'C'), ('E', 'F', 'G')])

        for letter, students in data.students_index:
            pos.next_letter()

            print(f"{pos.letter_merge()} : {pos.letter()}: {letter}")
            ws.merge_cells(pos.letter_merge())
            ws[pos.letter()] = letter
            ws[pos.letter()].style = 'indexletter'

            for s in sorted(list(students), key=lambda x: x.name):
                pos.next_student()
                ws[pos.pos(0)] = s.index_name
                ws[pos.pos(0)].style = 'indexstudent'
                ws[pos.pos(1)] = str(s.grade)
                ws[pos.pos(1)].style = 'indexstudent'
                ws[pos.pos(2)] = s.teacher.class_list_lookup.title()
                ws[pos.pos(2)].style = 'indexstudent'
                print(f"{pos.pos()}: {s.index_name:30} {s.grade} {s.teacher.class_list_lookup.title()}")
            print("")

        self.wb.save('output.xlsx')

class ExcelIndexPositioner:
    """
    Produces a position for writing students into the index
    using a two column structure
    """

    def __init__(self, page_height=62, page_buffer=4, columns=['A', 'E'], letter_start_index=4, student_start_index = 6):
        self.letter_start_index = letter_start_index
        self.student_start_index = student_start_index
        self.index = 2
        self.page_height = page_height
        self.page_buffer = page_buffer
        self.columns = columns
        self.column_index = 0
        self.page = 0

    def letter_merge(self):
        return f"{self.columns[self.column_index][0]}{self.index-1}:{self.columns[self.column_index][-1]}{self.index}"

    def letter(self):
        return f"{self.columns[self.column_index][0]}{self.index-1}"

    def pos(self, col=0):
        return f"{self.columns[self.column_index][col]}{self.index}"

    def next_letter(self):
        self.allocate_space(size=3, buffer=4, start_index=self.letter_start_index)

    def is_last_column(self):
        return self.column_index == len(self.columns) - 1

    def has_enough_space_in_column(self, buffer):
        limit = (self.page+1) * self.page_height - buffer
        print(f"  {self.index} < {limit}? {self.index < limit}")
        return self.index < limit

    def next_student(self):
        self.allocate_space(size=1, buffer=1, start_index=self.student_start_index)

    def allocate_space(self, size, buffer, start_index):
        if self.has_enough_space_in_column(buffer=buffer):
            # Enough space, use same column
            self.index += size
        elif self.is_last_column():
            # Move down a page
            self.page += 1
            self.column_index = 0
            self.index = self.page * self.page_height + start_index
        else:
            # Move over a column
            self.column_index = (self.column_index + 1) % len(self.columns)
            if self.page == 0:
                print(f"  moving to start index of {start_index}")
                self.index = start_index
            else:
                self.index = self.page * self.page_height + start_index

parser = argparse.ArgumentParser(prog='PROG', usage='%(prog)s [options]')
parser.add_argument('--pta-files', nargs='+', help='the PTA directory files')
parser.add_argument('--class-list', help='the class list file')

args = parser.parse_args()

class AllData:
    def __init__(self, class_lists, students):
        self.class_lists = class_lists
        self.students = students

        self.__update_class_list_data()
        self.__create_student_index()

    def __update_class_list_data(self):
        # Replace students in the class list with those from the parent information
        # since it contains more information
        for c in class_lists:
            pta_students = {s: s for s in students if s.teacher == c.teacher}

            class_students = sorted([pta_students[s] if s in pta_students else s for s in c.students])
            c.students = class_students

            # Replace the teacher with the data from the parent information since it has their full name
            c.teacher = class_students[0].teacher

    def __create_student_index(self):
        all_students = []
        for c in class_lists:
            all_students.extend(c.students)

        by_last_name_first_letter = lambda x: x.name[0]
        self.students_index = groupby(sorted(all_students, key=by_last_name_first_letter), key=by_last_name_first_letter)


class_lists = ClassListParser.parse_lists(args.class_list)
students = PtaParser.parse_pta_students(args.pta_files)

data = AllData(class_lists, students)

for output in [ ExcelOutput()]:
    for c in class_lists:
        output.print_class(c)
    output.finish(data)
