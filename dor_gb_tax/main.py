import argparse

from xlwt import Workbook
from xlrd import open_workbook


SHEETS = ['Leaves', 'Flowers', 'Immature Plants', 'Edibles', 'Concentrates',
          'Extracts', 'SkinHair Products', 'Other']
FIELDS = ['Date of Sale', 'Amount Sold', 'Exempt Amount Sold', 'Total Sales',
          'Tax-Exempt Sales', 'Taxable Sales']

QUARTERS = {1: 1, 2: 1, 3: 1, 4: 2, 5: 2, 6: 2, 7: 3, 8: 3, 9: 3, 10: 4, 11: 4,
            12: 4}
QUARTER_STRINGS = ['1st', '2nd', '3rd', '4th']


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
    # Initialize data preset
    data = {q: {s: [] for s in SHEETS} for q in QUARTER_STRINGS}
    for file_name in files:
        # Open the file
        wb = open_workbook(file_name)
        # Access the sheet
        sh = wb.sheet_by_index(0)
        for sheet in SHEETS:
            q = QUARTERS[file_name.split('-')[0]]
            x = SHEETS.index(sheet)
            # Build row of data for the month
            data[q][sheet].append([
                # Date of Sale (month-day-year)
                file_name.split('.')[0],
                # Amount Sold
                sh.cell_value(rowx=x + 1, colx=3),
                # Exempt Amount Sold
                sh.cell_value(rowx=x + 1, colx=1),
                # Total Sales
                float(sh.cell_value(rowx=x + 1, colx=2)) +
                float(sh.cell_value(rowx=x + 1, colx=4)),
                # Tax-Exempt Sales
                sh.cell_value(rowx=x + 1, colx=2),
                # Taxable Sales
                sh.cell_value(rowx=x + 1, colx=4),
            ])
    return(data)


def output_file(data):
    """
    With the dict of lists from process_data, create a sheet for each value in
    SHEETS. Then fill the data in from each row before creating the next sheet.
    """
    # Open new workbook
    wb = Workbook()
    for sheet_name in SHEETS:
        # Create new sheet with the appropriate label
        sh = wb.add_sheet(sheet_name)

        x, y = 0, 0
        # Write headers
        for field_name in FIELDS:
            sh.write(x, y, field_name)
            y += 1
        # Write data
        for row in data[sheet_name]:
            y = 0
            x += 1
            # Write cells
            for cell in row:
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
