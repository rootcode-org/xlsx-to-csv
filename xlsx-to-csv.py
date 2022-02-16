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


# Convert a zero-based integer index into a spreadsheet column identifier
# Not used but may be useful later
def index_to_column(index):
    output = ''
    while index >= 26:
        output = chr(ord('A') + index % 26) + output
        index = (index/26) - 1
    output = chr(ord('A') + index) + output
    return output


# Convert a spreadsheet column identifier to a zero-based integer index
# Not used but may be useful later
def column_to_index(column_name):
    value = 0
    for char in column_name:
        a = ord(char) - ord('A') + 1
        value = (value * 26) + a
    return value - 1


if __name__ == '__main__':

    if len(sys.argv) < 3:
        sys.exit(PURPOSE)

    worksheet_name = sys.argv[3] if len(sys.argv) > 3 else None
    xlsx_to_csv(sys.argv[1], sys.argv[2], worksheet_name)
