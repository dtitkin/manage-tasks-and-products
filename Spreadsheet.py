# модуль классов для работы API4 гугл таблиц
# в модуле реализованна минимальная часть работы с API
# достаточная для  "Управление продуктами и задачами"
# полная документация на API
# https://developers.google.com/sheets/api


from pprint import pprint
# import os

import httplib2
import apiclient.discovery
# import googleapiclient.errors
from oauth2client.service_account import ServiceAccountCredentials


def htmlColorToJSON(htmlColor):
    if htmlColor.startswith("#"):
        htmlColor = htmlColor[1:]
        rgb_red = {"red": int(htmlColor[0:2], 16) / 255.0}
        rgb_green = {"green": int(htmlColor[2:4], 16) / 255.0}
        rgb_blue = {"blue": int(htmlColor[4:6], 16) / 255.0}
    return rgb_red | rgb_green | rgb_blue


class SpreadsheetError(Exception):
    pass


class SpreadsheetNotSetError(SpreadsheetError):
    pass


class SheetNotSetError(SpreadsheetError):
    pass


class Spreadsheet():
    def __init__(self, jsonKeyFileName, debugMode=False):
        self.debugMode = debugMode
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(
            jsonKeyFileName,
            ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'])
        self.httpAuth = self.credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build(
            'sheets',
            'v4', http=self.httpAuth)
        self.spreadsheetId = None
        self.sheetId = None
        self.sheetTitle = None
        self.requests = []
        self.valueRanges = []

    def get_sheet_url(self):
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetId is None:
            raise SheetNotSetError()
        return 'https://docs.google.com/spreadsheets/d/'\
            + self.spreadsheetId + '/edit#gid='\
            + str(self.sheetId)

    def set_spreadsheet_byid(self, spreadsheetId):
        spreadsheet = self.service.spreadsheets().get(
            spreadsheetId=spreadsheetId).execute()
        if self.debugMode:
            pprint(spreadsheet)
        self.spreadsheetId = spreadsheet['spreadsheetId']
        self.sheetId = spreadsheet['sheets'][0]['properties']['sheetId']
        self.sheetTitle = spreadsheet['sheets'][0]['properties']['title']
        return spreadsheet

    # spreadsheets.batchUpdate и spreadsheets.values.batchUpdate
    def run_prepared(self, valueInputOption="USER_ENTERED"):
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        upd1Res = {'replies': []}
        upd2Res = {'responses': []}
        try:
            if len(self.requests) > 0:
                upd1Res = self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheetId,
                    body={"requests": self.requests}).execute()
                if self.debugMode:
                    pprint(upd1Res)
            if len(self.valueRanges) > 0:
                upd2Res = self.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self.spreadsheetId,
                    body={
                        "valueInputOption": valueInputOption,
                        "data": self.valueRanges}).execute()
                if self.debugMode:
                    pprint(upd2Res)
        finally:
            self.requests = []
            self.valueRanges = []
        return (upd1Res['replies'], upd2Res['responses'])

    # Converts string range to GridRange of current sheet; examples:
    # "A3:B4" -> {sheetId: id of current sheet, startRowIndex: 2,
    # endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
    # "A5:B"  -> {sheetId: id of current sheet, startRowIndex: 4,
    # startColumnIndex: 0, endColumnIndex: 2}
    def togrid_range(self, cellsRange):
        if self.sheetId is None:
            raise SheetNotSetError()
        if isinstance(cellsRange, str):
            startCell, endCell = cellsRange.split(":")[0:2]
            cellsRange = {}
            rangeAZ = range(ord('A'), ord('Z') + 1)
            start = 0
            abc = list(filter(lambda w: ord(w) in rangeAZ, startCell))
            num_startCell = startCell[len(abc):]
            if len(abc) > 1:
                start = 26 * (ord(abc[0]) - ord("A") + 1)
            cellsRange["startColumnIndex"] = start + ord(abc[-1]) - ord('A')

            start = 0
            abc = list(filter(lambda w: ord(w) in rangeAZ, endCell))
            num_endCell = startCell[len(abc):]
            if len(abc) > 1:
                start = 26 * (ord(abc[0]) - ord("A") + 1)
            cellsRange["endColumnIndex"] = start + ord(abc[-1]) - ord('A') + 1

            if len(num_startCell) > 0:
                cellsRange["startRowIndex"] = int(num_startCell) - 1
            if len(num_endCell) > 0:
                cellsRange["endRowIndex"] = int(num_endCell)
        cellsRange["sheetId"] = self.sheetId
        return cellsRange

    def prepare_setvalues(self, cellsRange, values, majorDimension="ROWS"):
        sheet_and_range = cellsRange.split("!")
        if len(sheet_and_range) == 2:
            self.sheetTitle = sheet_and_range[0]
            cellsRange = sheet_and_range[1]
        if self.sheetTitle is None:
            raise SheetNotSetError()
        self.valueRanges.append(
            {
                "range": self.sheetTitle + "!" + cellsRange,
                "majorDimension": majorDimension,
                "values": values})
        if self.debugMode:
            pprint(self.valueRanges)

    # formatJSON should be dict with userEnteredFormat
    # to be applied to each cell
    def prepare_setcells_format(self, cellsRange, formatJSON,
                                fields="userEnteredFormat"):
        self.requests.append(
            {
                "repeatCell": {
                    "range": self.togrid_range(cellsRange),
                    "cell": {"userEnteredFormat": formatJSON},
                    "fields": fields}})

    # formatsJSON should be list of lists of dicts with userEnteredFormat for
    # each cell in each row
    def prepare_setcells_many_format(self, cellsRange, formatsJSON,
                                     fields="userEnteredFormat"):

        rows = [{"values": [{"userEnteredFormat": cellFormat} for cellFormat in  rowFormats]} for rowFormats in formatsJSON]
        self.requests.append(
            {
                "updateCells": {"range": self.togrid_range(cellsRange),
                                "rows": rows,
                                "fields": fields}})

    def values_get(self, cellsRange, majorDimension="ROWS"):
        sheet_and_range = cellsRange.split("!")
        if len(sheet_and_range) == 2:
            self.sheetTitle = sheet_and_range[0]
            cellsRange = sheet_and_range[1]
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetTitle is None:
            raise SheetNotSetError()
        getRes = {}
        if len(cellsRange) > 0:
            getRes = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheetId,
                range=self.sheetTitle + "!" + cellsRange,
                majorDimension=majorDimension).execute()
            if self.debugMode:
                pprint(getRes)

        res = getRes.get("values")
        if not res:
            res = [[""]]
        return res

    def sh_get(self, cellsRange):
        ''' используем API spreadsheets
            для получения расширеннных данных из таблицы
        '''
        getRes = []
        sheet_and_range = cellsRange.split("!")
        if len(sheet_and_range) == 2:
            self.sheetTitle = sheet_and_range[0]
            cellsRange = sheet_and_range[1]
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetTitle is None:
            raise SheetNotSetError()
        if len(cellsRange) > 0:
            getRes = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheetId,
                ranges=self.sheetTitle + "!" + cellsRange,
                includeGridData=True).execute()
        if self.debugMode:
            pprint(getRes)
        return getRes

    def sheet_properies_get(self):
        '''для получения инфомрации о листе из раздела
            sheets properies
        '''
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetTitle is None:
            raise SheetNotSetError()

        # получаем данные листа
        getRes = self.service.spreadsheets().get(
            spreadsheetId=self.spreadsheetId,
            ranges=self.sheetTitle,
            includeGridData=False).execute()
        if self.debugMode:
            pprint(getRes)
        return getRes["sheets"][0]["properties"]

    def sheet_clear(self, cellsRange):
        '''для отчищает диапазон на листе
        '''
        sheet_and_range = cellsRange.split("!")
        if len(sheet_and_range) == 2:
            self.sheetTitle = sheet_and_range[0]
            cellsRange = sheet_and_range[1]
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetTitle is None:
            raise SheetNotSetError()
        getRes = None
        if len(cellsRange) > 0:
            getRes = self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheetId,
                range=self.sheetTitle + "!" + cellsRange).execute()
            if self.debugMode:
                pprint(getRes)
        return getRes

    def sheet_clear_all(self, name_sheet):
        ''' отчищает весь лист
        '''
        self.sheetTitle = name_sheet
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetTitle is None:
            raise SheetNotSetError()
        start_column = "A"
        max_column = self.max_column_get()
        resp = self.sheet_clear(f"{start_column}:{max_column}")
        return resp

    def read_color_cells(self, cellsRange):
        ''' возвращает список, элементы - строка
            строка - список столбцов, элементы - кортеж из трех словарей
            на каждый столбец
            1 - цвет фона
            2 - цвет шрифта
            3 - значение в ячейке
        '''
        return_list = []
        resp = self.sh_get(cellsRange)
        sh = resp["sheets"][0]
        data = sh["data"][0]
        row_data = data.get("rowData", [])
        n = 0
        for row in row_data:
            return_list.append([])
            row_lst = return_list[n]
            n += 1
            vls = row["values"]
            for column in vls:
                ef_format = column["effectiveFormat"]
                bg_color = ef_format["backgroundColor"]
                fnt_color = ef_format["textFormat"]["foregroundColor"]
                value = column["formattedValue"]
                row_lst.append((bg_color, fnt_color, value))
        return return_list

    def max_column_get(self):
        ''' возвращает буквенный адрес последней колонки активного листа- AB
        '''
        resp = self.sheet_properies_get()
        max_number = resp["gridProperties"]["columnCount"]
        rangeAZ = range(ord('A'), ord('Z') + 1)
        cont_symbol, num_symbol = divmod(max_number, 26)
        second_symbol = ""
        first_symbol = ""
        if cont_symbol == 0:
            first_symbol = chr(rangeAZ[max_number - 1])
        else:
            first_symbol = chr(rangeAZ[cont_symbol - 1])
            second_symbol = chr(rangeAZ[num_symbol - 1])
        return first_symbol + second_symbol

# Функции тестирования класса
'''  CREDENTIALS_FILE = 'creds.json'
TABLE_ID_TEST = "1ftW9wEf-zGMTtIe_YIfZZxgomRBJhkXjnLe9Q4eKFpM"


def test_set_spreadsheet_byid():
    print(CREDENTIALS_FILE)
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
    ss.set_spreadsheet_byid(TABLE_ID_TEST)
    print(ss.sheetId)


def test_togrid_range():
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
    ss.set_spreadsheet_byid(TABLE_ID_TEST)
    res = [ss.togrid_range("A3:B4"),
           ss.togrid_range("A5:B"),
           ss.togrid_range("A:B")]
    correctRes = [
        {
            "sheetId": ss.sheetId, "startRowIndex": 2, "endRowIndex": 4,
            "startColumnIndex": 0, "endColumnIndex": 2},
        {
            "sheetId": ss.sheetId, "startRowIndex": 4,
            "startColumnIndex": 0, "endColumnIndex": 2},
        {"sheetId": ss.sheetId, "startColumnIndex": 0, "endColumnIndex": 2}]
    print("Хорошо" if res == correctRes else "Плохо", res)


def test_cellformat():
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
    ss.set_spreadsheet_byid(TABLE_ID_TEST)
    ss.prepare_setcells_format(
        "B2:E7", {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"
                  })
    ss.run_prepared()


def test_cells_fieldarguments():
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
    ss.set_spreadsheet_byid(TABLE_ID_TEST)
    txt_field = "userEnteredFormat.textFormat,"
    txt_field = txt_field + "userEnteredFormat.horizontalAlignment"
    ss.prepare_setcells_format(
        "B1:F1",
        {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"},
        fields=txt_field)

    ss.prepare_setcells_format(
        "B1:F1",
        {"backgroundColor": htmlColorToJSON("#00CC00")},
        fields="userEnteredFormat.backgroundColor")
    ss.prepare_setcells_many_format(
        "C4:C4",
        [[{"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"}]],
        fields=txt_field)
    ss.prepare_setcells_many_format(
        "C4:C4",
        [[{"backgroundColor": htmlColorToJSON("#00CC00")}]],
        fields="userEnteredFormat.backgroundColor")
    pprint(ss.requests)
    ss.run_prepared()


def test_set_values():
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
    ss.set_spreadsheet_byid(TABLE_ID_TEST)
    values = [
        ["Название", "Количество", "Цена", "Дата", "Сумма"],
        ["Товар1", 10, 20, "15.01.2021", "=C2*D2"]]
    ss.prepare_setvalues("B1:F2", values)
    ss.run_prepared()


def test_get_values():
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
    ss.set_spreadsheet_byid(TABLE_ID_TEST)
    res = ss.values_get("B1:F2")


if __name__ == '__main__':
    test_set_spreadsheet_byid()
    test_togrid_range()
    test_cellformat()
    test_cells_fieldarguments()
    test_set_valuess()
    test_get_values()'''
