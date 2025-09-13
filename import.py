import openpyxl
import re
import sys

class Student:
    def __init__(self, name, grade, teacher):
        self.name = name
        self.grade = grade
        self.teacher = teacher

    def title(self):
        """
        Translate the student's name of the format "LAST, FIRST" to
        "First Last"
        """
        return re.sub(r"(.*?), (.*)", r"\2 \1", self.name).title()

    def __str__(self):
        return f"{self.title()} {self.grade if self.grade == 'K' else int(self.grade)} @ {self.teacher}"

class Teacher:
    def __init__(self, name):
        self.name = name

    def title(self):
        self.name.title()

    def __str__(self):
        return self.name

class Guardian:
    def __init__(self, name, email, phone=None, address=None):
        self.name = name
        self.email = email.lower()
        self.phone = phone
        self.address = address.title() if address else None

    def title(self):
        return self.name.title()

    def __str__(self):
        return f"{self.title()} {self.email} {self.phone} {self.address}"

students = []

for f in sys.argv[1:]:
    try:
        wb = openpyxl.load_workbook(f)

        sheet = wb.active

        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            student = Student(
                    name=row[0].value,
                    grade=row[1].value,
                    teacher=row[2].value)
            print(student)
            teacher = Teacher(name=row[3].value)

            has_phone = len(row) == 7
            has_address = len(row) == 10

            guardian = Guardian(
                    name=row[4].value,
                    email=row[5].value,
                    phone=row[6].value if has_phone else None,
                    address=row[7].value if has_address else None)
            print(guardian)
    except e:
        println(e)
