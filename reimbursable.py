import locale
import re
import os
from peewee import *

locale.setlocale(locale.LC_ALL, '')

# Config
sqlite_file = 'reimbursable.db'
output_file = 'output.xls'
sqlite_path = os.path.join(os.path.dirname(__file__), sqlite_file)
output_path = os.path.join(os.path.dirname(__file__), output_file)

# Input
input_amount_column = "D"
input_category_column = "F"
input_type_column = "E"
input_label_column = "H"
input_reimbursable_column = "J"

output_column = 1

db = SqliteDatabase(sqlite_path)

# Models
class Reimbursable(Model):
    name = CharField(index=True)

    class Meta:
        database = db

class Category(Model):
    name = CharField(index=True)

    class Meta:
        database = db

class Account(Model):
    name = CharField(index=True)

    class Meta:
        database = db

class Transaction(Model):
    reimbursable = ForeignKeyField(Reimbursable, related_name='transactions', null=True)
    category = ForeignKeyField(Category, related_name='transactions')
    account = ForeignKeyField(Account, related_name='transactions')
    date = DateField(index=True, null=True)
    amount = FloatField()
    row = IntegerField()

    class Meta:
        database = db
        # indexes = [(('date', 'row'), False)]

# Setup sqlite db if not found
if not os.path.isfile(sqlite_path):
    # Create tables
    Reimbursable.create_table()
    Category.create_table()
    Account.create_table()
    Transaction.create_table()

with db.transaction():
    for input_sheet in all_sheets():
        # Skip header
        input_row = 2

        account_name = input_sheet
        account = Account.get_or_create(name=account_name)

        while True:
            amount_cell = Cell(input_sheet, "{0}{1}".format(input_amount_column, input_row))
            category_cell = Cell(input_sheet, "{0}{1}".format(input_category_column, input_row))
            type_cell = Cell(input_sheet, "{0}{1}".format(input_type_column, input_row))
            label_cell = Cell(input_sheet, "{0}{1}".format(input_label_column, input_row))
            reimbursable_cell = Cell(input_sheet, "{0}{1}".format(input_reimbursable_column, input_row))

            # Stop when meet empty rows
            if amount_cell.is_empty():
                break

            amount = float(amount_cell.value)
            category_name = category_cell.value
            reimbursable_name = reimbursable_cell.value
            type = type_cell.value
            reimbursable = None

            # Make amount negative if credit
            if type == "credit":
                amount = -1*amount

            try:
                # If transaction exists in db, only update if
                # reimbursable is different
                transaction = Transaction.get(account=account, row=input_row)

            except Transaction.DoesNotExist:
                # If this transaction is reimbursable
                if reimbursable_name:
                    reimbursable = Reimbursable.get_or_create(name=reimbursable_name.title())

                category = Category.get_or_create(name=category_name)
                transaction = Transaction.create(reimbursable=reimbursable, category=category, account=account, amount=amount, row=input_row)

            if transaction.reimbursable:
                if not transaction.reimbursable.name == reimbursable_name:
                    transaction.reimbursable = reimbursable
                    transaction.save

            print "Imported transaction #" + str(transaction.id)

            # Next row
            input_row += 1

# Remove old output
try:
    os.remove(output_path)
except OSError:
    pass

# Create new workbook
active_wkbk(new_wkbk())

# Now display it all
for reimbursable in Reimbursable.select():
    output_sheet = reimbursable.name
    new_sheet(output_sheet)

    totals = {}

    for transaction in reimbursable.transactions:
        category = transaction.category

        if category.name not in totals:
            totals[category.name] = 0.00

        totals[category.name] += transaction.amount

    row = 2
    grand_total = 0

    for category_name, total in iter(sorted(totals.iteritems())):
        # Skip categories with 0
        if total == 0:
            continue

        category_cell = Cell(output_sheet, row, 1)
        total_cell = Cell(output_sheet, row, 2)

        category_cell.value = category_name
        total_cell.value = locale.currency(total)

        grand_total += total
        row += 1

    grand_total_label_cell = Cell(output_sheet, 1, 1)
    grand_total_label_cell.value = "Total"
    grand_total_label_cell.font.bold = True

    grand_total_cell = Cell(output_sheet, 1, 2)
    grand_total_cell.value = locale.currency(grand_total)
    grand_total_cell.font.bold = True

    # Autofit sheet cells
    autofit(output_sheet)

# Remove default sheets
for sheet in all_sheets():
    if re.match(r'Sheet', sheet):
        remove_sheet(sheet)

# Save output
save(output_path)
