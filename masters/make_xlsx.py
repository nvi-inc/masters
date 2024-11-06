import re

from pathlib import Path
from datetime import datetime
from string import ascii_uppercase
from copy import copy

from openpyxl import load_workbook
from openpyxl.utils import rows_from_range
from openpyxl.worksheet.worksheet import Worksheet

from masters import app, get_master_file


class MasterText:
    """
    Class to read master text file and generate a xlsx file
    """
    def __init__(self, path: Path):
        """
        Initialize class
        :param path: Path of text file
        """
        self.path = path
        template = 'master-template.xlsx' if app.args.master else 'int-template.xlsx'
        self.wb = load_workbook(Path(app.config['folder'], template))

    def __enter__(self):
        self.file = open(self.path, 'r')
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.file.close()

    def copy_header(self, sheet: Worksheet) -> None:
        """
        copy header of template to new sheet
        :param sheet: Worksheet to be updated
        :return: None
        """
        for row in rows_from_range('A1:AS1'):
            for cell in row:
                old_cell = self.template[cell]
                sheet[cell].value = old_cell.value
                if old_cell.has_style:
                    new_cell = sheet[cell]
                    new_cell.font = copy(old_cell.font)
                    new_cell.alignment = copy(old_cell.alignment)
                    new_cell.border = copy(old_cell.border)
                    new_cell.fill = copy(old_cell.fill)
                    new_cell.number_format = old_cell.number_format

    @staticmethod
    def hm2m(duration: str) -> float:
        h, _, m = [a.strip() for a in duration.partition(':')]
        minutes = float(h) if h else 0
        return minutes + (float(m) if m else 0)

    def process_master(self) -> Path:
        """
        Read text file and generate xlsx file
        :return:
        """
        skd = [i + j for i in ('', 'A') for j in ascii_uppercase][6:36]
        rmv = skd[::-1]
        lines = [line for line in self.file if line.startswith('|')]
        # Access the active sheet
        sheet = self.wb.active
        sheet.title = f"{app.args.year} MS"
        for row, line in enumerate(lines, 2):
            ses = [info.strip() for info in line.split('|')[1:]]
            sheet[f'A{row}'], sheet[f'B{row}'] = ses[0], ses[2].upper()
            sheet[f'C{row}'] = datetime.strptime(ses[1], '%Y%m%d')
            sheet[f'E{row}'] = datetime.strptime(ses[4], '%H:%M').time()
            sheet[f'F{row}'] = self.hm2m(ses[5])
            sheet[f'AK{row}'], sheet[f'AL{row}'] = ses[7], ses[8]
            sheet[f'AM{row}'] = datetime.strptime(ses[9], '%Y%m%d') if ses[9].isdigit() else ses[9]
            sheet[f'AO{row}'], sheet[f'AP{row}'] = ses[10], ses[11]
            scheduled, _, removed = ses[6].partition(' -')
            for col in skd:  # Empty field for each stations
                sheet[f'{col}{row}'] = ''
            for col, sta in zip(skd, re.findall(r'..', scheduled)):
                sheet[f'{col}{row}'] = f'{sta}1G-'
            sheet[f'{col}{row}'] = f'{sta}1G'
            end = ''
            for col, sta in zip(rmv, re.findall(r'..', removed)):
                sheet[f'{col}{row}'] = f'{sta}1G{end}'
                end = '-'

        # Clean empty lines
        sheet.delete_rows(row + 1, 500)
        # Save file
        self.wb.save(xlsx := get_master_file())
        return xlsx

    def process_intensive(self) -> Path:
        """
        Read text file and generate xlsx file
        :return:
        """
        skd = [i for i in ascii_uppercase[12:22]]
        rmv = skd[::-1]
        lines = [line for line in self.file if line.startswith('|')]
        # Access the active sheet
        sheet = self.wb.active
        sheet.title = f"{app.args.year} INT"
        for row, line in enumerate(lines, 2):
            ses = [info.strip() for info in line.split('|')[1:]]
            sheet[f'A{row}'], sheet[f'C{row}'] = ses[0], ses[2].upper()
            sheet[f'E{row}'] = datetime.strptime(ses[1], '%Y%m%d')
            sheet[f'I{row}'] = datetime.strptime(ses[4], '%H:%M').time()
            sheet[f'K{row}'] = self.hm2m(ses[5])
            sheet[f'X{row}'], sheet[f'Z{row}'] = ses[7], ses[8]
            sheet[f'AB{row}'] = datetime.strptime(ses[9], '%Y%m%d') if ses[9].isdigit() else ses[9]
            sheet[f'AF{row}'], sheet[f'AH{row}'] = ses[10], ses[11]

            scheduled, _, removed = ses[6].partition(' -')
            for col in skd:  # Empty field for each stations
                sheet[f'{col}{row}'] = ''
            for col, sta in zip(skd, re.findall(r'..', scheduled)):
                sheet[f'{col}{row}'] = f'{sta}-'
            sheet[f'{col}{row}'] = f'{sta}'
            end = ''
            for col, sta in zip(rmv, re.findall(r'..', removed)):
                sheet[f'{col}{row}'] = f'{sta}{end}'
                end = '-'
            for col in ['B', 'D', 'F', 'H', 'J', 'L', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI']:
                sheet[f'{col}{row}'] = '|'

        # Clean empty lines
        sheet.delete_rows(row + 1, 1500)
        # Save file
        self.wb.save(xlsx := get_master_file())
        return xlsx


    def process(self):
        print(f'Reading {self.path}')
        xlsx = self.process_master() if app.args.master else self.process_intensive()
        print(f'Created {xlsx}')

def main():
    import argparse

    parser = argparse.ArgumentParser(description='Generate master file')
    parser.add_argument('-c', '--config', help='config file', required=True)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-master', action='store_true')
    group.add_argument('-intensives', action='store_true')
    parser.add_argument('year', help='master file year', type=int)

    app.init(parser.parse_args())
    with MasterText(get_master_file(extension='txt')) as master:
        master.process()


if __name__ == '__main__':

    import sys
    sys.exit(main())
