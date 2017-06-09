import argparse

from xlwt import Workbook
from xlrd import open_workbook


SHEETS = ['Leaves', 'Flowers', 'Immature Plants', 'Edibles', 'Concentrates',
          'Extracts', 'SkinHair Products', 'Other']
FIELDS = ['Date of Sale', 'Amount Sold', 'Exempt Amount Sold', 'Total Sales',
          'Tax-Exempt Sales', 'Taxable Sales']


def main():
    """
    A program written to take the files expoted by greenbits and reassemble
    them for import into the oregon online revenue system.
    """
    files = set_args()
    data = process_data(files)
    output_file(data)


def process_data(files):
    """
    For each file name in the list, open and read the file. Create a dictionary
    object of lists for each sheet name in SHEETS. For each file opened, append
    each row's data to the appropriate sheet.
    """
    data = {x: [] for x in SHEETS}
    for file_name in files:
        wb = open_workbook(file_name)
        sh = wb.sheet_by_index(0)
        for sheet in SHEETS:
            x = SHEETS.index(sheet)
            data[sheet].append([
                file_name.split('.')[0],
                sh.cell_value(rowx=x + 1, colx=3),
                sh.cell_value(rowx=x + 1, colx=1),
                sh.cell_value(rowx=x + 1, colx=2) + sh.cell_value(
                    rowx=x + 1, colx=4),
                sh.cell_value(rowx=x + 1, colx=2),
                sh.cell_value(rowx=x + 1, colx=4),
            ])
    return(data)


def output_file(data):
    """
    With the dict of lists from process_data, create a sheet for each value in
    SHEETS. Then fill the data in from each row before creating the next sheet.
    """
    wb = Workbook()
    for sheet_name in SHEETS:
        sh = wb.add_sheet(sheet_name)

        x, y = 0, 0
        for field_name in FIELDS:
            sh.write(x, y, field_name)
            y += 1
        for col_data in data[sheet_name]:
            y = 0
            x += 1
            for cell in col_data:
                sh.write(x, y, cell)
                y += 1

    wb.save('output.xls')


def set_args():
    """ Init args with argparse """
    parser = argparse.ArgumentParser(
        description='Turn monthly reports into importable file for taxes. '
        'Please Note, filenames MUST be in a \'day/month/year.xls\' format!')
    parser.add_argument('files', metavar='FILE_LIST', type=str, nargs='+',
                        help='One or more files that are imported')
    return parser.parse_args().files


if __name__ == "__main__":
    main()
