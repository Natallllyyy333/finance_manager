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

    for t in transactions:
        if t['type'] == 'income':
            analysis['total_income'] += t['amount']
        else:
            analysis['total_expenses'] += t['amount']
            analysis['categories'][t['category']] += t['amount']
            analysis['categories'][t['date']] += t['amount']
            analysis['daily_spending'][t['date']] += t['amount']
            analysis['savings'] = analysis['total_income'] - \
                analysis['total_expenses']
            return analysis


def generate_ai_recomendations(analysis):
    """Generate AI recommendations based on financial analysis."""
    recomendations = []
    biggest_category = max(analysis['categories'].items(), key=lambda x: x[1])
    recomendations.append(
        f"Consider reducing spending in {biggest_category[0]} category, which accounts for {biggest_category[1]:.2f} of your expenses.")

    savings_rate = (analysis['savings'] / analysis['total_income']
                    ) * 100 if analysis['total_income'] > 0 else 0
    if savings_rate < 20:
        recomendations.append(
            f"Your savings rate is low at {savings_rate:.2f}%. Consider increasing your savings up to 20% of your income by reducing discretionary spending.")

        expensive_day = max(
            analysis['daily_spending'].items(), key=lambda x: x[1])
        recomendations.append(
            f"Your most expensive day was {expensive_day[0]} with a total spending of {expensive_day[1]:.2f}. Consider reviewing your expenses on that day. ")
        return recomendations


def visualize_data(analysis):
    """Visualize financial data using matplotlib."""
    # Making a bar chart for spending by category 10x5 inches size
    plt.figure(figsize=(10, 5))
    # categories = list(analysis['categories'].keys())
    # amounts = list(analysis['categories'].values())

    # plt.figure(figsize=(10, 6))
    # plt.bar(categories, amounts, color='skyblue')
    # plt.xlabel('Categories')
    # plt.ylabel('Amount')
    # plt.title('Spending by Category')
    # plt.xticks(rotation=45)
    # plt.tight_layout()
    # plt.show()


sa = gspread.service_account(filename='creds.json')
sh = sa.open("Personal Finances")
wks = sh.worksheet(f"{MONTH}")
rows = hsbcFin(file)
for row in rows:
    wks.insert_row([row[0], row[1], row[2], row[3], row[4]], 8)
    time.sleep(2)  # Sleep to avoid hitting API limits
