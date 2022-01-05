from datetime import datetime
from collections import namedtuple

from openpyxl.cell import Cell
from openpyxl import load_workbook


DisciplineType = namedtuple('Discipline', 'number name study_group lecturer room')


class ExcelApi:
    SHEET_NAME = '30'

    def __init__(self, filename):
        self.disciplines = {}
        self.wb = load_workbook(filename, data_only=True)
        self.sheet = self.wb.get_sheet_by_name(self.SHEET_NAME)

        self.get_dates_with_disciplines()

    def get_dates_with_disciplines(self):
        for i in range(1, 1200):
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

        lecturer1 = self.sheet.cell(row=row, column=col + 5).value
        if lecturer1 is not None:
            room_value = self.sheet.cell(row=row, column=col + 7).value
            room = room_value if room_value is not None else self.sheet.cell(row=4, column=col + 7).value
            disciplines.append(DisciplineType(number, name, study_group, lecturer1, room))

        lecturer2 = self.sheet.cell(row=row, column=col + 6).value
        if lecturer2 is not None:
            room_value = self.sheet.cell(row=row, column=col + 8).value
            room = room_value if room_value is not None else self.sheet.cell(row=4, column=col + 8).value
            disciplines.append(DisciplineType(number, name, study_group, lecturer2, room))

    def add_disciplines(self, date_cell: Cell) -> list:
        row, col = date_cell.row, date_cell.column + 4

        disciplines = []
        while col < 300:
            self.add_discipline(disciplines, '1-2', row, col)
            self.add_discipline(disciplines, '3-4', row + 1, col)
            self.add_discipline(disciplines, '5-6', row + 2, col)
            self.add_discipline(disciplines, '7-8', row + 3, col)
            col += 9

        return disciplines

    def get_discipline(self, date, lecturer) -> list:
        res_disciplines = []

        if self.disciplines.get(date) is None:
            return res_disciplines

        for discipline in self.disciplines[date]['disciplines']:
            if discipline.lecturer == lecturer:
                res_disciplines.append(discipline)
        return res_disciplines
