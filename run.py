# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
from collections import defaultdict
from datetime import datetime

DAILY_NORMS = {
    'Rent': 30,
    'Gym': 2,
    'Groceries': 15,
    'Transport': 8,
    'Entertainment': 5,
    'Utilities': 3,
    'Shopping': 10,
}
print("\n" + " PERSONAL FINANCE ANALYZER ".center(80, "="))
MONTH = input(
    "Enter the month (e.g., 'March', 'April', 'May'): ").strip().capitalize()
FILE = f"hsbc_{MONTH.lower()}.csv"

# file = f"hsbc_{MONTH}.csv"


def load_transactions(filename):
    # with open(file, mode='r', encoding='utf-8') as csv_file:
    """Load and categorize transactions with daily tracking"""
    # csv_reader = csv.reader(csv_file)
    transactions = []
    daily_categories = defaultdict(lambda: defaultdict(float))
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            for row in csv.reader(f):
                if len(row) < 5:
                    continue
                try:
                    amount = float(row[2])
                    category = categorize(row[1])
                    date = row[0]

                    transactions.append({
                        'date': date,
                        'desc': row[1][:30],
                        'amount': amount,
                        'type': 'income' if row[4] == 'Credit' else 'expense',
                        'category': category
                    })
                    if row[4] != 'Credit':
                        daily_categories[date][category] += amount
                except ValueError:
                    continue  # Skip rows with invalid data
    except FileNotFoundError:
        print(f"\nError: File '{filename}' not found")
        exit()
    return transactions, daily_categories


def categorize(description):
    """Categorize transaction based on description."""
    desc = description.lower()
    categories = {
        'Rent': ['rent', 'monthly rent'],
        'Groceries': ['supermarket', 'grocery', 'food'],
        'Dining': ['restaurant', 'cafe', 'coffee'],
        'Transport': ['bus', 'train', 'taxi', 'uber'],
        'Entertainment': ['movie', 'netflix', 'concert'],
        'Other': []
    }
    for cat, terms in categories.items():
        if any(term in desc for term in terms):
            return cat
    return 'Other'


def analyze(transactions, daily_categories):
    """Perform financial analysis with daily tracking"""
    analysis = {
        'income': 0, 'expenses': 0, 'categories': defaultdict(float),
        'month': MONTH, 'daily_categories': daily_categories,
        'days_count': len(daily_categories) or 1
    }
    for t in transactions:
        if t['type'] == 'income':
            analysis['income'] += t['amount']
        else:
            analysis['expenses'] += t['amount']
            analysis['categories'][t['category']] += t['amount']
    analysis['savings'] = analysis['income'] - analysis['expenses']
    return analysis


def tetrminal_visualization(data):
    """Visualize financial data in terminal."""
    # Header
    print(f"\n +{data['month'].upper()} FINANCIAL OVERVIEW ".center(80, "="))
    # Summary bars
    print(f"\nIncome: {data['income']:10.2f}€ [" + "■" *
          int(data['income'] / max(data['income'], 1) * 30) + "]")
    print("fExpenses:{data['expenses']:10.2f}€ [" + "■" *
          int(data['expenses'] / max(data['income'], 1) * 30) + "]")
    print(f"Savings: {data['savings']:10.2f}€ [" + "■" *
          int(data['savings'] / max(data['income'], 1) * 30) + "]")

    # Categories breakdown
    print("\n EXPENSE CATEGORIES ".center(80, '-'))
    for cat, amount in sorted(data['categories'].items(), key=lambda x: x[1], reverse=True):
        pct = amount / data['expenses'] * 100 if data['expenses'] > 0 else 0
        print(f"{cat:12}:<12 {amount:7.2f}€ " +
              "■" * int(pct / 5) + f" {pct:.1f}%")


def generate_daily_recommendations(data):
    """Generate daily category-specific recommendations."""
    recs = []
    if data['income'] <= 0:
        return ["No income data - cannot generate recommendations."]

     # 1. Savings rate recommendation
    savings_rate = (data['savings'] / data['income'] * 100)
    if savings_rate < 20:
        recs.append(f"Aim for 20% savings (current: {savings_rate:.1f}%)")
        # 2. Daily category norms analysis
    for day, categories in data['daily_categories'].items():
        for cat, amount in categories.items():
            if cat in DAILY_NORMS and amount > DAILY_NORMS[cat]*1.2:
                recs.append(f"On {day}, {cat} spent {amount:.2f}€"
                            f"norm: {DAILY_NORMS[cat]}€")
        # 3.Monthly category averages vs daily norms
    for cat, norm in DAILY_NORMS.items():
        if cat in data['categories']:
            avg_daily = data['categories'][cat] / data['days_count']
            if avg_daily > norm * 1.1:  # 10% over norm
                recs.append(
                    f"Reduce daily{cat} spending "
                    f"(current: {avg_daily:.2f}€, norm: {norm}€)")

    # Ensure minimum recommendations
    if len(recs) < 3:
        recs.extend([
            "Plan meals weekly to reduce grocery costs",
            "Use public transport more frequently",
            "Limit dining out to 2-3 times a week"
        ])
        return recs[:5]  # Return only top 5 recommendations


def main():
    transactions, daily_categories = load_transactions(FILE)
    if not transactions:
        print(f"No transactions found")
        return
    data = analyze(transactions, daily_categories)
    tetrminal_visualization(data)

    # Recommendations
    print("\n DAILY SPENDING RECOMMENDATIONS ".center(80, '='))
    for i, rec in enumerate(generate_daily_recommendations(data), 1):
        print(f"{i}. {rec}")

    # Optional Google Sheets update
    try:
        gs = gspread.service_account('creds.json')
        sheet = gs.open("Personal Finances").worksheet(MONTH)
        sheet.clear()
        sheet.append_row(["Date", "Description", "Amount", "Type", "Category"])
        rows = [
            [t['date'], t['desc'], t['amount'], t['type'], t['category']]
            for t in transactions
        ]
        sheet.insert_rows(rows)
        print("\nData saved to Google Sheets")
    except Exception:
        pass  # Skip if no credentials


if __name__ == "__main__":
    main()

    # for row in csv_reader:
    # if len(row) < 5:
    #     continue
    # try:
    #     # date = datetime(row[0])
    #     date = row[0]
    #     name = row[1]
    #     amount = float(row[2])
    #     transaction_type = 'income' if row[4] == 'Credit' else 'expense'
    #     category = categorize_transaction(name)
    #     if name == 'Monthly Rent Payment':
    #         category = 'rent'
    #     elif name == 'Gym Membership':
    #         category = 'gym'
    #     elif name == 'Gym Membership':
    #         category = 'gym'
    #     elif name == 'Gym Membership':
    #         category = 'gym'
    #     elif name == 'Gym Membership':
    #         category = 'gym'
    #     transactions.append({
    #         'date': date,
    #         'name': name,
    #         'amount': amount,
    #         'type': transaction_type,
    #         'category': category
    #     })

    # transaction = ((date, name, amount, category, transation_type))

    # transactions.append(transaction)
    #     except Exception as e:
    #         print(f"Error processing row {row}: {e}")
    # return transactions


# def categorize_transaction(name):
#     name = name.lower()
#     categories = {
#         'rent': ['rent', 'monthly rent'],
#         'gym': ['gym', 'fitness'],
#         'groceries': ['supermarket', 'grocery'],
#         'utilities': ['electricity', 'water', 'gas'],
#         'entertainment': ['cinema', 'restaurant', 'bar'],
#         'transport': ['bus', 'train', 'taxi'],
#         'shopping': ['clothing', 'electronics']
#     }
#     for category, keywords in categories.items():
#         if any(keyword in name.lower() for keyword in keywords):
#             return category
#     return 'other'


# def analyze_finances(transactions):

#     analysis = {
#         'total_income': 0,
#         'total_expenses': 0,
#         'categories': defaultdict(float),
#         'daily_spending': defaultdict(float)
#     }

#     for t in transactions:
#         if t['type'] == 'income':
#             analysis['total_income'] += t['amount']
#         else:
#             analysis['total_expenses'] += t['amount']
#         analysis['categories'][t['category']] += t['amount']
#         # analysis['categories'][t['date']] += t['amount']
#         analysis['daily_spending'][t['date']] += t['amount']
#         analysis['savings'] = analysis['total_income'] - \
#             analysis['total_expenses']
#     return analysis


# def generate_ai_recommendations(analysis):
#     """Generate AI recommendations based on financial analysis."""
#     recommendations = []
#     biggest_category = max(analysis['categories'].items(), key=lambda x: x[1])
#     recommendations.append(
#         f"Consider reducing spending in {biggest_category[0]} category, which accounts for {biggest_category[1]:.2f} of your expenses.")

#     savings_rate = (analysis['savings'] / analysis['total_income']
#                     ) * 100 if analysis['total_income'] > 0 else 0
#     if savings_rate < 20:
#         recommendations.append(
#             f"Your savings rate is low at {savings_rate:.2f}%. Consider increasing your savings up to 20% of your income by reducing discretionary spending.")

#         expensive_day = max(
#             analysis['daily_spending'].items(), key=lambda x: x[1])
#         recommendations.append(
#             f"Your most expensive day was {expensive_day[0]} with a total spending of {expensive_day[1]:.2f}. Consider reviewing your expenses on that day. ")
#     return recommendations


# def visualize_data(analysis):
#     """Visualize financial data using matplotlib."""
#     # Making a bar chart for spending by category 10x5 inches size
#     plt.figure(figsize=(10, 5))
#     # Circle diagram for income categories
#     if analysis['categories']:
#         # First subplot for categories(1 row, 2 columns, first plot)
#         plt.subplot(1, 2, 1)
#         plt.pie(analysis['categories'].values(), labels=analysis['categories'].keys(),
#                 autopct='%1.1f%%')
#         plt.title('Spending by Category')
#         # Daily spending chart
#         dates = sorted(analysis['daily_spending'].keys())
#         amounts = [analysis['daily_spending'][d] for d in dates]
#         # Second subplot for daily spending(1 row, 2 columns, second plot)
#         plt.subplot(1, 2, 2)
#         # Line chart for daily spending with markers
#         plt.plot(dates, amounts, marker='o')
#         plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
#         plt.title('Daily Spending')  # Title for daily spending chart
#         plt.tight_layout()  # Adjust layout to prevent overlap
#         plt.show()


# def update_google_sheet(worksheet, transactions):
#     """Update Google Sheet with transactions."""
#     if len(worksheet.get_all_values()) > 1:
#         worksheet.delete_rows(2, len(worksheet.get_all_values()))
#     for t in transactions:
#         row = [
#             t['date'],
#             t['name'],
#             t['amount'],
#             t['type'],
#             t['category']
#         ]
#         worksheet.append_row(row)

#     time.sleep(2)


# def main():
#     print(f"Processing transactions for {MONTH.capitalize()}...")
#     transactions = hsbcFin(file)

#     print(f"Analyzing finances for {MONTH.capitalize()}...")
#     analysis = analyze_finances(transactions)
#     print(f"Financial analysis for {MONTH.capitalize()}")
#     print(f"Total Income: {analysis['total_income']:.2f}€")
#     print(f"Total Expenses: {analysis['total_expenses']:.2f}€")
#     print(f"Savings: {analysis['savings']:.2f}€")

#     print("Expenditure by Category: ")
#     for category, amount in analysis['categories'].items():
#         print(f" - {category}: {amount:.2f}€")

#     print("\n == Recommendations == ")
#     recommendations = generate_ai_recommendations(analysis)
#     for i, recommendation in enumerate(recommendations, 1):
#         print(f"{i}.{recommendation}")
#     print("\nVisualizing data...")
#     visualize_data(analysis)
#     print("\nUpdating Google Sheet...")
#     try:
#         sa = gspread.service_account(filename='creds.json')
#         sh = sa.open("Personal Finances")
#         try:
#             wks = sh.worksheet(MONTH)
#         except gspread.WorksheetNotFound:
#             print(f"Worksheet for {MONTH} not found. Creating a new one.")
#             wks = sh.add_worksheet(title=MONTH, rows="100", cols="10")
#             headers = ["Date", "Description", "Amount", "Type", "Category"]
#             wks.append_row(headers)
#         update_google_sheet(wks, transactions)
#         print("Google Sheet updated successfully.")
#     except Exception as e:
#         print(f"Error updating Google Sheet: {e}")


# if __name__ == "__main__":
#     main()

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


# sa = gspread.service_account(filename='creds.json')
# sh = sa.open("Personal Finances")
# wks = sh.worksheet(f"{MONTH}")
# rows = hsbcFin(file)
# for row in rows:
#     wks.insert_row([row[0], row[1], row[2], row[3], row[4]], 8)
#     time.sleep(2)  # Sleep to avoid hitting API limits
