import gspread

gc = gspread.service_account('service_account_credentials.json')

spreadsheet = gc.open('Employee Sales')

while True:
    ws1_name = input('Please enter the name of the first worksheet: ')
    rng1_addr = input('Please specify the target range of cells on the first worksheet: ')

    ws2_name = input('Please enter the name of the second worksheet: ')
    rng2_addr = input('Please specify the target range of cells on the second worksheet: ')

    ws1 = spreadsheet.worksheet(ws1_name)
    ws2 = spreadsheet.worksheet(ws2_name)

    rng1 = ws1.get(rng1_addr)
    rng2 = ws2.get(rng2_addr)

    num_rows = len(rng1)
    num_cols = len(rng1[0])

    if num_rows == len(rng2) and num_cols == len(rng2[0]):
        break
    else:
        print('Please select two ranges with identical dimensions.')

headers = rng1[0]

diffs = []

for i in range(1, num_rows):
    for j in range(num_cols):
        cell1 = rng1[i][j]
        cell2 = rng2[i][j]

        if cell1 != cell2:
            if cell1.isnumeric():
                cell1 = float(cell1)

            if cell2.isnumeric():
                cell2 = float(cell2)

            diff = [i, headers[j], cell1, cell2]

            diffs.append(diff)

diffs_ws = spreadsheet.add_worksheet(title='Diffs', rows=100, cols=100)

diffs_rng = 'A2:D' + str(len(diffs) + 1)

diffs_ws.update(diffs_rng, diffs)

diffs_ws.update('A1:E1', [['Row', 'Column', 'Value 1', 'Value 2', 'Delta']])

formula_rng = 'E2:E' + str(len(diffs) + 1)

formulas = []

for idx, row in enumerate(diffs):
    current_row = str(idx + 2)

    if type(row[2]) == float and type(row[3]) == float:
        cell1_addr = 'C' + current_row
        cell2_addr = 'D' + current_row

        formula = '=ABS(' + cell1_addr + '-' + cell2_addr + ')/AVERAGE(' + cell1_addr + ':' + cell2_addr + ')'

        formulas.append([formula])
    else:
        formulas.append([])

diffs_ws.update(formula_rng, formulas, raw=False)

delta_style = {
    "textFormat": {
        "foregroundColor": {
            "red": 1.0,
            "green": 0.0,
            "blue": 0.0
        }
    },
    "numberFormat": {
        "type": "PERCENT"
    }
}

diffs_ws.format(formula_rng, delta_style)






