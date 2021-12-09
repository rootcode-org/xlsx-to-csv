# Copyright is waived. No warranty is provided. Unrestricted use and modification is permitted.

import sys

try:
    import openpyxl
except ImportError:
    sys.exit('Requires openpyxl module; try "pip install openpyxl"')


PURPOSE = '''\
Convert an XLSX file to CSV

xlsx-to-csv.py <input_file> <output_file> [<worksheet>]

where,
   <input_file>    Path to input xlsx file
   <output_file>   Path to output csv file
   <worksheet>     Name of worksheet to convert; if not specified first worksheet is converted
'''


def xlsx_to_csv(input_file, output_file, worksheet_name):

    # Open the workbook
    try:
        book = openpyxl.load_workbook(input_file)
    except Exception as e:
        sys.exit('ERROR: Error opening XLSX file')

    # Select the named worksheet; if no name is specified select the first worksheet
    if worksheet_name is None:
        worksheet_name = book.get_sheet_names()[0]
    sheet = book.get_sheet_by_name(worksheet_name)

    # Get the expected width of the array
    width = sheet.max_column

    # Convert cell values to CSV
    csv_rows = []
    for row in sheet.rows:
        csv_cells = []
        for cell in row:
            value = cell.value
            if value is None:
                value = ''
            csv_cells.append(str(value))
        row_string = ','.join(csv_cells)
        csv_rows.append(row_string)

    # Write output string to CSV file
    with open(output_file, 'w') as f:
        f.write('\n'.join(csv_rows))


if __name__ == '__main__':

    if len(sys.argv) < 3:
        sys.exit(PURPOSE)

    worksheet_name = sys.argv[3] if len(sys.argv) > 3 else None
    xlsx_to_csv(sys.argv[1], sys.argv[2], worksheet_name)
