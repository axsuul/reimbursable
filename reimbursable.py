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
    weight = IntegerField(default=0)

    class Meta:
        database = db

class Category(Model):
    name = CharField(index=True)

    def description(self):
        description = self.name;

        if self.percent() < 1:
            description += " (" + str(self.percentage()) + "%)"

        return description

    def percentage(self):
        return int(self.percent()*100)

    def percent(self):
        if re.match(r'Gas|Auto|Fuel', self.name):
            return 0.5

        if re.match(r'Food|Restaurants|Dining', self.name):
            return 0.5

        return 1

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

    def calculate_amount(self):
        return self.category.percent()*self.amount

    class Meta:
        database = db

# Fresh db
try:
    os.remove(sqlite_path)

except OSError:
    pass

# Create tables
Reimbursable.create_table()
Category.create_table()
Account.create_table()
Transaction.create_table()

# Create a nonreimbursable for aynthing that
# can't be reimbursed
nonreimbursable = Reimbursable.create(name="Non-expensable", weight=100)

# Output helper
def output_transactions(sheet,  label, transactions, column):
    totals = {}

    for transaction in transactions:
        category = transaction.category

        if category.name not in totals:
            totals[category.name] = { 'amount': 0.00, 'category': category }

        totals[category.name]['amount'] += transaction.calculate_amount()

    row = 2
    grand_total = 0

    for category_name, data in iter(sorted(totals.iteritems())):
        total = data['amount']
        category = data['category']

        # Skip categories with 0
        if total == 0:
            continue

        category_cell = Cell(output_sheet, row, column)
        total_cell = Cell(output_sheet, row, column + 1)
        reimbursable_cell = Cell(output_sheet, row, column + 2)

        category_cell.value = category.description()
        total_cell.value = locale.currency(total)
        reimbursable_cell.value = reimbursable.name

        grand_total += total
        row += 1

    Cell(output_sheet, 1, column).value = label
    Cell(output_sheet, 1, column + 1).value = locale.currency(grand_total)

def account_transactions(sheet, transactions, reimbursable, row, column):
    totals = {}
    grand_total = 0

    for transaction in transactions:
        category = transaction.category

        if category.name not in totals:
            totals[category.name] = { 'amount': 0.00, 'category': category }

        totals[category.name]['amount'] += transaction.calculate_amount()

    for category_name, data in iter(sorted(totals.iteritems())):
        total = data['amount']
        category = data['category']

        # Skip categories with 0
        if total == 0:
            continue

        category_cell = Cell(output_sheet, row, column)
        total_cell = Cell(output_sheet, row, column + 1)
        reimbursable_cell = Cell(output_sheet, row, column + 2)

        category_cell.value = category.description()
        total_cell.value = locale.currency(total)
        reimbursable_cell.value = reimbursable.name

        grand_total += total
        row += 1

    Cell(output_sheet, row, column).value = "Total"
    Cell(output_sheet, row, column).font.bold = True
    Cell(output_sheet, row, column + 1).value = locale.currency(grand_total)
    Cell(output_sheet, row, column + 1).font.bold = True
    Cell(output_sheet, row, column + 2).value = reimbursable.name
    Cell(output_sheet, row, column + 2).font.bold = True

    return { 'total': grand_total, 'row': row }

with db.transaction():
    for input_sheet in all_sheets():
        account_name = input_sheet
        account = Account.get_or_create(name=account_name)

        for input_row in list(set(all_cells(input_sheet).row)):
            # Skip header
            if input_row < 2:
                continue

            amount_cell = Cell(input_sheet, "{0}{1}".format(input_amount_column, input_row))
            category_cell = Cell(input_sheet, "{0}{1}".format(input_category_column, input_row))
            type_cell = Cell(input_sheet, "{0}{1}".format(input_type_column, input_row))
            reimbursable_cell = Cell(input_sheet, "{0}{1}".format(input_reimbursable_column, input_row))

            # Stop when meet empty rows
            if amount_cell.is_empty():
                break

            amount = float(amount_cell.value)
            category_name = category_cell.value
            type = type_cell.value
            reimbursable_name = reimbursable_cell.value
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
                # Only consider debits for non-reimbursable
                elif type == "debit":
                    reimbursable = nonreimbursable

                category = Category.get_or_create(name=category_name)
                transaction = Transaction.create(reimbursable=reimbursable, category=category, account=account, amount=amount, row=input_row)

            if transaction.reimbursable:
                if not transaction.reimbursable.name == reimbursable_name:
                    transaction.reimbursable = reimbursable
                    transaction.save

            print "Imported transaction #" + str(transaction.id)

# Remove old output
try:
    os.remove(output_path)
except OSError:
    pass

# Create new workbook
active_wkbk(new_wkbk())

# Now display it all
output_sheet = "Totals"
new_sheet(output_sheet)

column = 1

for account in Account.select():
    # Build headers
    Cell(output_sheet, 1, column).value = account.name
    Cell(output_sheet, 1, column + 1).value = "Charge"
    Cell(output_sheet, 1, column + 2).value = "Type"

    reimbursables = Reimbursable.select().order_by(Reimbursable.weight.asc())
    reimbursable_totals = {}
    row = 2

    for reimbursable in reimbursables:
        transactions = account.transactions.select().where(Transaction.reimbursable == reimbursable)

        # Output
        data = account_transactions(output_sheet, transactions, reimbursable, row, column)

        reimbursable_totals[reimbursable.name] = data['total']

        # Skip row for next
        row = data['row'] + 2

    # Leave column gap
    column += 4

    # Format everythang
    header_range = CellRange(output_sheet, "A1:Z1")
    header_range.color = "black"
    header_range.font.color = "white"
    header_range.font.bold = True

    # Autofit sheet cells
    autofit(output_sheet)

# Remove default sheets
for sheet in all_sheets():
    if re.match(r'Sheet', sheet):
        remove_sheet(sheet)

# Save output
save(output_path)
