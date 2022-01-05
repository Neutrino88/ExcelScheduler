from datetime import datetime
from collections import namedtuple

from openpyxl.cell import Cell
from openpyxl import load_workbook


DisciplineType = namedtuple('Discipline', 'number name study_group lecturers rooms')


class ExcelApi:
    SHEET_NAME = '30'

    def __init__(self, filename):
        self.disciplines = {}
        self.wb = load_workbook(filename, data_only=True)
        self.sheet = self.wb.get_sheet_by_name(self.SHEET_NAME)

        self.get_dates_with_disciplines()

    def get_dates_with_disciplines(self):
        for i in range(1, 600):
            cell = self.sheet.cell(row=i, column=1)

            if isinstance(cell.value, datetime):
                self.disciplines[cell.value.date()] = {
                    'cell': cell,
                    'disciplines': self.add_disciplines(cell),
                }

    def add_discipline(self, disciplines, number, row, col):
        if self.sheet.cell(row=row, column=col).value is None:
            return

        name = self.sheet.cell(row=row, column=col).value
        study_group = self.sheet.cell(row=4, column=col).value

        lecturers = [self.sheet.cell(row=row, column=col + 5).value, self.sheet.cell(row=row, column=col + 6).value]

        room1_value = self.sheet.cell(row=row, column=col + 7).value
        room1 = room1_value if room1_value is not None else self.sheet.cell(row=4, column=col + 7).value

        room2_value = self.sheet.cell(row=row, column=col + 8).value
        room2 = room2_value if room2_value is not None else self.sheet.cell(row=4, column=col + 8).value

        disciplines.append(DisciplineType(number, name, study_group, lecturers, (room1, room2)))

    def add_disciplines(self, date_cell: Cell) -> list:
        row, col = date_cell.row, date_cell.column + 4

        disciplines = []
        while col < 300:
            self.add_discipline(disciplines, 1, row, col)
            self.add_discipline(disciplines, 2, row + 1, col)
            self.add_discipline(disciplines, 3, row + 2, col)
            self.add_discipline(disciplines, 4, row + 3, col)
            col += 9

        return disciplines

    def get_discipline(self, date, lecturer):
        res_disciplines = []

        if self.disciplines.get(date) is None:
            return res_disciplines

        for discipline in self.disciplines[date]['disciplines']:
            for disc_lecturer in discipline.lecturers:
                if disc_lecturer == lecturer:
                    res_disciplines.append(discipline)
        return res_disciplines
