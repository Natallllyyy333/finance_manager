# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
MONTH = input(
    "Enter the month (e.g., 'March', 'April', 'May'): ").strip().lower()

file = f"hsbc_{MONTH}.csv"

with open(file, mode='r') as csv_file:
    csv_reader = csv.reader(csv_file)
    for row in csv_reader:
        date = row[0]
        name = row[1]
        amount = float(row[2])
        category = 'other'
        transaction = ((date, name, amount, category))
        print(transaction)

sa = gspread.service_account(filename='creds.json')
sh = sa.open("Personal Finances")
