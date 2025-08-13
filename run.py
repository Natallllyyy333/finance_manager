# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
import time
import matplotlib.pyplot as plt
from collections import defaultdict
from datetime import datetime
MONTH = input(
    "Enter the month (e.g., 'March', 'April', 'May'): ").strip().lower()

file = f"hsbc_{MONTH}.csv"


def hsbcFin(file):
    with open(file, mode='r', encoding='utf-8') as csv_file:
        csv_reader = csv.reader(csv_file)
        transactions = []
        for row in csv_reader:
            if len(row) < 5:
                continue
            try:
                date = datetime.strptime(row[0], '%d%m%Y')
                date = row[0]
                name = row[1]
                amount = float(row[2])
                transaction_type = 'income' if row[4] == 'Credit' else 'expense'
                category = categorize_transaction(name)
                transactions.apppend({
                    'date': date,
                    'name': name,
                    'amount': amount,
                    'type': transaction_type,
                    'category': category
                })
                if name == 'Monthly Rent Payment':
                    category = 'rent'
                elif name == 'Gym Membership':
                    category = 'gym'
                # transaction = ((date, name, amount, category, transation_type))

                # transactions.append(transaction)
            except Exception as e:
                print(f"Error processing row {row}: {e}")
        return transactions


def categorize_transaction(name):
    name = name.lower()
    categories = {
        'rent': ['rent', 'monthly rent'],
        'gym': ['gym', 'fitness'],
        'groceries': ['supermarket', 'grocery'],
        'utilities': ['electricity', 'water', 'gas'],
        'entertainment': ['cinema', 'restaurant', 'bar'],
        'transport': ['bus', 'train', 'taxi'],
        'shopping': ['clothing', 'electronics']
    }
    for category, keywords in categories.items():
        if any(keyword in name.lower() for keyword in keywords):
            return category
    return 'other'


def analyze_finances(transactions):

    analysis = {
        'total_income': 0,
        'total_expenses': 0,
        'categories': defaultdict(float),
        'daily_spending': defaultdict(float)
    }


sa = gspread.service_account(filename='creds.json')
sh = sa.open("Personal Finances")
wks = sh.worksheet(f"{MONTH}")
rows = hsbcFin(file)
for row in rows:
    wks.insert_row([row[0], row[1], row[2], row[3], row[4]], 8)
    time.sleep(2)  # Sleep to avoid hitting API limits
