import locale
import re
import sqlite3

locale.setlocale(locale.LC_ALL, '')

# Input
inputAmountColumn = "D"
inputCategoryColumn = "F"
inputTypeColumn = "E"
inputLabelColumn = "H"
inputReimbursableColumn = "I"

# Output
outputSheet = "Totals"

# Clear sheet
clear_sheet(outputSheet)

outputColumn = 1
grandTotals = {}
reimbursementTotals = {}

# Our own data structure that stores everything
data = {}

# Helper to output totals
def output_totals(name, totals, row, column):
    total = 0.00
    headerRow = row
    row += 1    # make room for header

    for category, amount in iter(sorted(totals.iteritems())):
        # Skip categories with 0
        if amount == 0:
            continue

        categoryCell = Cell(outputSheet, row, column)
        amountCell = Cell(outputSheet, row, column + 1)

        categoryCell.value = category
        amountCell.value = locale.currency(amount)

        total += amount
        row += 1

    # Header
    nameCell = Cell(outputSheet, headerRow, column)
    nameCell.font.bold = True
    nameCell.value = name
    totalCell = Cell(outputSheet, headerRow, column + 1)
    totalCell.font.bold = True
    totalCell.value = locale.currency(total)

    return row

outputRow = 1

for inputSheet in all_sheets():
    # Skip output sheet
    if inputSheet == outputSheet:
        continue

    totals = {}

    inputRow = 2

    while True:
        amountCell = Cell(inputSheet, "{0}{1}".format(inputAmountColumn, inputRow))
        categoryCell = Cell(inputSheet, "{0}{1}".format(inputCategoryColumn, inputRow))
        typeCell = Cell(inputSheet, "{0}{1}".format(inputTypeColumn, inputRow))
        labelCell = Cell(inputSheet, "{0}{1}".format(inputLabelColumn, inputRow))
        reimbursableCell = Cell(inputSheet, "{0}{1}".format(inputReimbursableColumn, inputRow))

        # Stop when meet empty rows
        if amountCell.is_empty():
            break

        amount = float(amountCell.value)
        category = categoryCell.value
        type = typeCell.value
        reimbursable = False

        if reimbursableCell.value:
            reimbursable = True

        if category not in totals:
            totals[category] = 0.00

        if category not in grandTotals:
            grandTotals[category] = 0.00

        if reimbursable:
            if type == "credit":
                amount = -1*amount

            totals[category] += amount
            grandTotals[category] += amount

        # Next row
        inputRow += 1

    outputRow = output_totals(inputSheet, totals, outputRow, 5)

    # Skip for the next sheet
    outputRow += 1



output_totals("Totals", grandTotals, 1, 1)