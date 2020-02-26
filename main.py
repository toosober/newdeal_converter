import sys
from collections import namedtuple, UserList

import xlrd


class Way(UserList):
    def __init__(self, *args, name):
        super().__init__(*args)
        self.name = name


class Account:
    def __init__(self, type_):
        self._type = type_
        self.resources = Way(name=Parser.RESOURCES)
        self.uses = Way(name=Parser.USES)


Record = namedtuple('Record', 'name code values children')


class Parser:
    CODE = 'коды'
    RESOURCES = 'Ресурсы'
    USES = 'Использование'
    TOTAL = 'Всего'
    TITLE_COLUMN = 2
    NON_EMPTY_TYPES = (xlrd.XL_CELL_TEXT, xlrd.XL_CELL_NUMBER)

    def __init__(self):
        self._accounts = []
        self._years = None
        self._years_len = None
        self._state = Parser.look_for_year_string
        self._current_way = None

    def process(self, row):
        self._state(self, row)

    def look_for_year_string(self, row):
        values = [cell.value for cell in row]
        if Parser.CODE in values:
            ind = values.index(Parser.CODE)
            self._years = [cell.value for cell in row[ind+1:] if cell.ctype in Parser.NON_EMPTY_TYPES]
            self._years_len = len(self._years)
            self._state = Parser.look_for_accounts_type

    def look_for_accounts_type(self, row):
        title_cell: xlrd.sheet.Cell = row[Parser.TITLE_COLUMN]
        if title_cell.ctype == xlrd.XL_CELL_TEXT:
            account = Account(title_cell.value)
            self._current_way = account.resources
            self._accounts.append(account)
            self._state = Parser.look_for_way_type

    def look_for_way_type(self, row):
        account = self._accounts[-1]
        title_cell: xlrd.sheet.Cell = row[Parser.TITLE_COLUMN]
        if self._current_way.name == title_cell.value:
            self._state = Parser.look_for_account_record

    def look_for_account_record(self, row):
        account = self._accounts[-1]
        name = row[0].value
        code = row[1].value
        values = row[2:2+self._years_len]
        new_record = Record(name, code, values, [])
        if not values[0].value:  # пустые значения, не интересная строка
            return
        try:
            prefix = code[:-1]
            assert len(prefix) > 0
            record = [record for record in self._current_way if record.code == prefix].pop()
            record.children.append(new_record)
        except:
            self._current_way.append(new_record)

        if name == Parser.TOTAL:
            if self._current_way.name == Parser.USES:
                self._current_way = None
                self._state = Parser.look_for_accounts_type
                return
            elif self._current_way.name == Parser.RESOURCES:
                self._current_way = account.uses
                self._state = Parser.look_for_way_type
            else:
                raise Exception()


def main(path_):
    workbook = xlrd.open_workbook(path_, formatting_info=True)
    sheet = workbook.sheet_by_index(0)
    parser = Parser()
    for row in sheet.get_rows():
        parser.process(row)
    print(parser._accounts)
    print(parser._years)


if __name__ == "__main__":
    path = sys.argv[1]
    main(path)

