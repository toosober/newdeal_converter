import sys
import json
import dataclass_factory
from typing import List, Dict, Tuple, Callable, NewType, Union
from dataclasses import dataclass, field
import xlrd


@dataclass
class Record:
    name: str
    code: str
    values: List[Tuple[int, Union[float, None]]] = field(default_factory=list)
    children: List['Record'] = field(default_factory=list)


@dataclass
class Way:
    name: str
    records: List[Record] = field(default_factory=list)


@dataclass
class Account:
    type: str
    resources: Way = field(default_factory=lambda: Way(Parser.RESOURCES))
    uses: Way = field(default_factory=lambda: Way(Parser.USES))


@dataclass
class Document:
    accounts: List[Account] = field(default_factory=list)


StateProcessor = NewType('StateProcessor', Callable[['Parser', str], Union['StateProcessor', None]])


def look_for_year_string(parser, row):
    values = [cell.value for cell in row]
    if Parser.CODE in values:
        ind = values.index(Parser.CODE)
        raw_years = (cell.value for cell in row[ind+1:] if cell.ctype in Parser.NON_EMPTY_TYPES)
        years = (str(raw_year).split(' ')[0] for raw_year in raw_years)  # чистим от сносок
        parser._years = [int(float(year)) for year in years]
        parser._years_len = len(parser._years)
        return look_for_accounts_type


def look_for_accounts_type(parser, row):
    title_cell: xlrd.sheet.Cell = row[Parser.TITLE_COLUMN]
    if title_cell.ctype == xlrd.XL_CELL_TEXT:
        account = Account(title_cell.value)
        parser._current_way = account.resources
        parser.document.accounts.append(account)
        return look_for_way_type


def look_for_way_type(parser, row):
    title_cell: xlrd.sheet.Cell = row[Parser.TITLE_COLUMN]
    if parser._current_way.name == title_cell.value:
        return look_for_account_record


def look_for_account_record(parser, row):
    account = parser.document.accounts[-1]
    name: str = row[0].value
    code = row[1].value
    raw_values = (cell.value for cell in row[2:])
    values = [(year, value) for year, value in zip(parser._years, raw_values)]
    if not any(value[1] for value in values):  # нет значений, неинтересная строка
        return

    new_record = Record(name, code, values)

    # вложенная запись может быть либо с расширяемым кодом, например D.3 -> D.39
    # либо начинаться с "в том числе" проверим сначала одно, потом второе
    try:
        prefix = code[:-1]
        assert len(prefix) > 0
        record = [record for record in parser._current_way.records if record.code == prefix].pop()
        record.children.append(new_record)
    except:
        INCLUDING = "в том числе "
        if name.startswith(INCLUDING):
            new_record.name = new_record.name[len(INCLUDING):]
            last_record: Record = parser._current_way.records[-1]
            last_record.children.append(new_record)
        else:
            parser._current_way.records.append(new_record)

    if name == Parser.TOTAL:
        if parser._current_way.name == Parser.USES:
            parser._current_way = None
            return look_for_accounts_type
            return
        elif parser._current_way.name == Parser.RESOURCES:
            parser._current_way = account.uses
            return look_for_way_type
        else:
            raise Exception()


@dataclass
class Parser:
    document: Document = field(default_factory=Document)
    _years: List[str] = None
    _state: StateProcessor = look_for_year_string
    _current_way = None
    _years_len = None

    CODE = 'коды'
    RESOURCES = 'Ресурсы'
    USES = 'Использование'
    TOTAL = 'Всего'
    TITLE_COLUMN = 2
    NON_EMPTY_TYPES = (xlrd.XL_CELL_TEXT, xlrd.XL_CELL_NUMBER)

    def process(self, row):
        new_state = self._state(self, row)
        if new_state:
            self._state = new_state


def main(path_):
    workbook = xlrd.open_workbook(path_, formatting_info=True)
    sheet = workbook.sheet_by_index(0)
    parser = Parser()
    for row in sheet.get_rows():
        parser.process(row)
    factory = dataclass_factory.Factory()
    dct = factory.dump(parser.document)
    print(json.dumps(dct, ensure_ascii=False))


if __name__ == "__main__":
    path = sys.argv[1]
    main(path)

