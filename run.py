# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
import time
MONTH = input(
    "Enter the month (e.g., 'March', 'April', 'May'): ").strip().lower()

file = f"hsbc_{MONTH}.csv"


def hsbcFin(file):
    with open(file, mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        transactions = []
        for row in csv_reader:
            date = row[0]
            name = row[1]
            amount = float(row[2])
            category = 'other'
            if name == 'Monthly Rent Payment':
                category = 'rent'
            elif name == 'Gym Membership':
                category = 'gym'
            transaction = ((date, name, amount, category))

            transactions.append(transaction)
        return transactions


sa = gspread.service_account(filename='creds.json')
sh = sa.open("Personal Finances")
wks = sh.worksheet(f"{MONTH}")
rows = hsbcFin(file)
for row in rows:
    wks.insert_row([row[0], row[1], row[2], row[3]], 8)
    time.sleep(2)  # Sleep to avoid hitting API limits
