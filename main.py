import json

import openpyxl as pyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class ExcelIO:
    worksheet: Worksheet
    workbook: Workbook
    alphabet = list(chr(n) for n in range(65, 91))

    def __init__(self, file_path=None):
        try:
            self.workbook = pyxl.load_workbook(file_path)
        finally:
            pass
        self.name = str
        self.worksheet = ...

    def set_workbook(self, file_path):
        self.workbook = pyxl.load_workbook(file_path)
        self.name = file_path
        self.worksheet = self.workbook[self.workbook.sheetnames[0]]

    @staticmethod
    def get_cell(x: int, y: int, letter: str = '') -> str:

        while x >= 1:
            if x <= 26:
                letter += ExcelIO.alphabet[x - 1]
                x = 0
            else:
                letter += ExcelIO.alphabet[int(x) // 26 - 1]
                x %= 26
        return letter + str(y)

    def extract_headers(self, sheet_number):
        self.worksheet = self.workbook[self.workbook.sheetnames[sheet_number]]
        export_col = 1
        excel_data = []
        val = self.worksheet[self.get_cell(export_col, 1)].value
        while val is not None:
            excel_data.append(val)
            export_col += 1
            val = self.worksheet[self.get_cell(export_col, 1)].value
        return excel_data

    def extract_data(self, sheet_number):
        self.worksheet = self.workbook[self.workbook.sheetnames[sheet_number]]
        excel_data = []
        export_row = 1
        export_col = 1
        while self.worksheet[self.get_cell(1, export_row)].value is not None:
            export_row += 1
        while self.worksheet[self.get_cell(export_col, 1)].value is not None:
            export_col += 1
        print(export_col, export_row)
        for r in range(2, export_row):
            excel_data.append([])
            for c in range(1, export_col):
                val = self.worksheet[self.get_cell(c, r)].value
                try:
                    val = int(val)
                except ValueError:
                    pass
                except TypeError:
                    pass
                finally:
                    excel_data[r-2].append(val)
        return excel_data


class SqlIO:
    def __init__(self):
        pass


class JsonIO:

    def __init__(self, file_path):
        self.file_path = file_path
        self.file = open(self.file_path, 'w', encoding="utf-8", errors="replace")
        self.data = None
        self.file.close()

    def add_entries(self, txt, keys):

        with open(self.file_path, 'r', encoding="utf-8", errors="replace") as self.file:
            self.data = []
        print(txt)
        for n, txt_input in enumerate(txt):
            print(txt_input)
            entry = {}
            for index, key in enumerate(keys):
                print(key, txt_input[index])
                entry[key] = txt_input[index]
            self.data.append(entry)

        self.file = open(self.file_path, 'w', encoding="utf-8", errors="replace")
        self.file.write(json.dumps(self.data, sort_keys=False, indent=4))
        self.file.close()

    def clear_entries(self):
        with open(self.file_path, 'w', encoding="utf-8", errors="replace") as self.file:
            self.file.write(json.dumps([], sort_keys=False, indent=4))

    def get_entries(self):
        try:
            with open(self.file_path, 'r') as self.file:
                self.data = json.load(self.file)
                print("DATA:", self.data)
        except FileNotFoundError:
            print("NOT FOUND?")
        print(self.data)
        return self.data


json_data = {0: JsonIO("output/airplanes.json"),
             1: JsonIO("output/airports.json"),
             2: JsonIO("output/locations.json"),
             3: JsonIO("output/persons.json"),
             4: JsonIO("output/flights.json"),
             5: JsonIO("output/routes.json")}

if __name__ == '__main__':
    spreadsheet = ExcelIO('data.xlsx')
    for n in range(6):
        json_data[n].add_entries(spreadsheet.extract_data(n), spreadsheet.extract_headers(n))
