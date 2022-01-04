import asyncio
from datetime import datetime
from collections import namedtuple

from openpyxl.cell import Cell
from openpyxl import load_workbook


DisciplineType = namedtuple('Discipline', 'number name topic lesson type hours lecturers rooms')


class ExcelApi:
    SHEET_NAME = '30'

    def __init__(self, filename):
        self.disciplines = {}
        self.wb = load_workbook(filename, data_only=True)
        self.sheet = self.wb.get_sheet_by_name(self.SHEET_NAME)

        asyncio.create_task(self.get_dates_with_disciplines())

    async def get_dates_with_disciplines(self):
        for i in range(1, 2000):
            cell = self.sheet.cell(row=i, column=1)

            if isinstance(cell.value, datetime):
                self.disciplines[cell.value.date()] = {
                    'cell': cell,
                    'disciplines': await self.add_disciplines(cell),
                }

    async def add_discipline(self, disciplines, number, row, col):
        if self.sheet.cell(row=row, column=col).value is None:
            return

        name = self.sheet.cell(row=row, column=col).value
        topic = self.sheet.cell(row=row, column=col + 1).value
        lesson = self.sheet.cell(row=row, column=col + 2).value
        _type = self.sheet.cell(row=row, column=col + 3).value
        hours = self.sheet.cell(row=row, column=col + 4).value
        lecturers = [self.sheet.cell(row=row, column=col + 5).value, self.sheet.cell(row=row, column=col + 6).value]
        rooms = (self.sheet.cell(row=row, column=col + 7).value, self.sheet.cell(row=row, column=col + 8).value)

        disciplines.append(DisciplineType(number, name, topic, lesson, _type, hours, lecturers, rooms))

    async def add_disciplines(self, date_cell: Cell) -> list:
        row, col = date_cell.row, date_cell.column + 4

        disciplines = []
        while col < 300:
            await self.add_discipline(disciplines, 1, row, col)
            await self.add_discipline(disciplines, 2, row + 1, col)
            await self.add_discipline(disciplines, 3, row + 2, col)
            await self.add_discipline(disciplines, 4, row + 3, col)
            col += 9

        return disciplines

    async def get_discipline(self, date, lecturer):
        res_disciplines = []
        for discipline in self.disciplines[date]:
            for disc_lecturer in discipline.lecturers:
                if disc_lecturer == lecturer:
                    res_disciplines.append(discipline)
        return res_disciplines
