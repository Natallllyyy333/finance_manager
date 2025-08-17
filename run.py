# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
from collections import defaultdict
import time
from datetime import datetime

DAILY_NORMS = {
    'Rent': 1200,
    'Gym': 45,
    'Groceries': 90,
    'Transport': 8,
    'Entertainment': 5,
    'Utilities': 20,
    'Shopping': 100,
}
print("\n" + " PERSONAL FINANCE ANALYZER ".center(80, "="))
MONTH = input(
    "Enter the month (e.g., 'March', 'April', 'May'): ").strip().lower()
FILE = f"hsbc_{MONTH}.csv"

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
        'Income': ['salary', 'bonus', 'income'],
        'Rent': ['rent', 'monthly rent'],
        'Groceries': ['supermarket', 'grocery', 'food'],
        'Dining': ['restaurant', 'cafe', 'coffee'],
        'Transport': ['bus', 'train', 'taxi', 'uber'],
        'Entertainment': ['movie', 'netflix', 'concert'],
        'Utilities': ['electricity', 'water', 'gas', 'internet', 'phone'],
        'Gym': ['gym', 'Gym Membership' 'fitness', 'yoga'],
        'Shopping': ['clothing', 'electronics', 'shopping', 'Supermarket'],
        'Health': ['pharmacy', 'doctor', 'health'],
        'Insurance': ['insurance', 'health insurance', 'car insurance'],
        'Education': ['tuition', 'books', 'courses', 'course'],
        'Travel': ['flight', 'hotel', 'travel', 'airline'],
        'Savings': ['savings', 'investment', 'stocks'],
        'Bank Fees': ['bank fee', 'atm fee', 'service charge'],
        'Charity': ['donation', 'charity', 'fundraiser'],
        'Car': ['car', 'vehicle', 'fuel', 'maintenance'],
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
    print(
        "\n" + f" {data['month'].upper()} FINANCIAL OVERVIEW ".center(80, "="))
    # Summary bars
    expense_rate = (data['expenses'] / data['income']
                    * 100) if data['income'] > 0 else 0
    savings_rate = (data['savings'] / data['income']
                    * 100) if data['income'] > 0 else 0
    print(f"Income: {data['income']:10.2f}€ [" + "■" *
          int(data['income'] / max(data['income'], 1) * 30) + "]" + " " + "100%")
    print(f"\nExpenses: {data['expenses']:10.2f}€ [" + "■" *
          int(data['expenses'] / max(data['income'], 1) * 30) + "]" + " " + f"{expense_rate:.1f}%")
    print(f"Savings: {data['savings']:10.2f}€ [" + "■" *
          int(data['savings'] / max(data['income'], 1) * 30) + "]" + " " + f"{savings_rate:.1f}%")

    # Categories breakdown
    print("\n" + f" EXPENSE CATEGORIES ".center(80, '-'))
    for cat, amount in sorted(data['categories'].items(), key=lambda x: x[1], reverse=True):
        pct = amount / data['expenses'] * 100 if data['expenses'] > 0 else 0
        print(f"{cat:<12}{amount:7.2f}€ " +
              "■" * int(pct / 5) + f" {pct:.1f}%")


def generate_daily_recommendations(data):
    """Generate daily category-specific recommendations."""
    recs = []
    if not data or 'income' not in data:
        return ["No financial data available for recommendations."]

    if data['income'] <= 0:
        return ["No income data - cannot generate recommendations."]
    else:

        # 1. Savings rate recommendation
        expense_rate = (data['expenses'] / data['income'] * 100)
        savings_rate = (data['savings'] / data['income'] * 100)
        if savings_rate < 20:
            recs.append(f"Aim for 20% savings (current: {savings_rate:.1f}%)")

            # 2. Daily category norms analysis
        for day, categories in data['daily_categories'].items():
            for cat, amount in categories.items():
                if cat in DAILY_NORMS and amount > DAILY_NORMS[cat]*1.2:
                    recs.append(f"On {day}, {cat} spent {amount:.2f}€. "
                                f"Norm: {DAILY_NORMS[cat]}€")

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


# def main():
#     transactions, daily_categories = load_transactions(FILE)
#     if not transactions:
#         print(f"No transactions found")
#         return
#     data = analyze(transactions, daily_categories)
#     tetrminal_visualization(data)

#     # Recommendations
#     print("\n" + f" DAILY SPENDING RECOMMENDATIONS ".center(80, '='))
#     for i, rec in enumerate(generate_daily_recommendations(data), 1):
#         print(f"{i}. {rec}")

#     # Optional Google Sheets update
#     try:
#         # Authenticate and open Google Sheets
#         gs = gspread.service_account('creds.json')
#         sh = gs.open("Personal Finances")
#         worksheet = None
#         try:
#             worksheet = sh.worksheet(MONTH)
#             print(f"\nUpdating Google Sheets for {MONTH}...")
#         except gspread.WorksheetNotFound:
#             print(f"\nWorksheet for {MONTH} not found. Creating a new one...")

#             worksheet = sh.add_worksheet(title=MONTH, rows="100", cols="5")
#             print(f"New worksheet '{MONTH}' created.")

#          # Now we're sure worksheet exists
#         if worksheet is None:
#             raise Exception("Failed to create or access worksheet")

#         # Clear existing data (keep headers if they exist)
#         all_values = worksheet.get_all_values()

#         # Clear existing data and set headers
#         # Clear existing data (except headers)
#         if len(worksheet.get_all_values()) > 1:
#             worksheet.delete_rows(2, len(worksheet.get_all_values()))
#             print(f"Worksheet '{MONTH}' already exists. Using it...")
#             # worksheet = sh.worksheet(MONTH)
#             # else:
#             #     raise e

#             # Set headers
#         headers = ["Date", "Description", "Amount", "Type", "Category"]
#         worksheet.append_row(headers)
#         # Append transactions
#         rows = []
#         for t in transactions:
#             rows.append(
#                 [t['date'], t['desc'], str(
#                     t['amount']), t['type'], t['category']]
#             )

#             # Append all rows at once
#         if rows:
#             # worksheet.append_rows(rows)
#             for i in range(0, len(rows), 100):
#                 batch = rows[i:i+100]
#                 worksheet.append_rows(batch)
#                 time.sleep(1)  # Brief pause between batches
#                 print("\nData saved to Google Sheets")
#         else:
#             print("\nNo transactions to save to Google Sheets")
#     except gspread.exceptions.APIError as e:
#         print(f"\nGoogle Sheets API Error: {str(e)}")
#         print("Possible solutions:")
#         print("1. Check if the sheet name contains invalid characters")
#         print("2. Verify your service account has proper permissions")
#     except Exception as e:
#         print(f"\nGeneral error: {str(e)}")
def main():
    transactions, daily_categories = load_transactions(FILE)
    if not transactions:
        print(f"No transactions found")
        return
    data = analyze(transactions, daily_categories)
    tetrminal_visualization(data)

    # Recommendations
    print("\n" + f" DAILY SPENDING RECOMMENDATIONS ".center(80, '='))
    for i, rec in enumerate(generate_daily_recommendations(data), 1):
        print(f"{i}. {rec}")

    # Optional Google Sheets update
    try:
        # Authenticate and open Google Sheets
        gs = gspread.service_account('creds.json')
        sh = gs.open("Personal Finances")

        # Check if worksheet exists
        worksheet = None
        try:
            worksheet = sh.worksheet(MONTH)
            print(f"\nWorksheet '{MONTH}' found. Updating...")
        except gspread.WorksheetNotFound:
            print(f"\nWorksheet for {MONTH} not found. Creating a new one...")
            # First check if we've reached the sheet limit (max 200 sheets)
            if len(sh.worksheets()) >= 200:
                raise Exception("Maximum number of sheets (200) reached")

            # Check if sheet exists but with different case (e.g. "march" vs "March")
            existing_sheets = [ws.title for ws in sh.worksheets()]
            if MONTH.lower() in [sheet.lower() for sheet in existing_sheets]:
                # Find the existing sheet with case-insensitive match
                for sheet in sh.worksheets():
                    if sheet.title.lower() == MONTH.lower():
                        worksheet = sheet
                        print(
                            f"Using existing worksheet '{sheet.title}' (case difference)")
                        break
            else:
                # Create new worksheet with unique name if needed
                try:
                    worksheet = sh.add_worksheet(
                        title=MONTH, rows="100", cols="20")
                    print(f"New worksheet '{MONTH}' created successfully.")
                except gspread.exceptions.APIError as e:
                    if "already exists" in str(e):
                        # If we get here, it means the sheet exists but wasn't found earlier
                        worksheet = sh.worksheet(MONTH)
                        print(f"Worksheet '{MONTH}' exists. Using it.")
                    else:
                        raise e

        if worksheet is None:
            raise Exception("Failed to access or create worksheet")

        # Clear existing data (keep headers)
        all_values = worksheet.get_all_values()
        if len(all_values) > 1:
            worksheet.delete_rows(7, len(all_values)+7)

            headers = ["Date", "Description", "Amount", "Type", "Category"]
            worksheet.insert_row(headers, 7)

        # Add transactions in batches
        if transactions:
            batch_size = 1
            for i in range(0, len(transactions), batch_size):
                batch = [
                    [
                        t['date'],
                        t['desc'][:50],
                        t['amount'],
                        t['type'],
                        t['category']
                    ] for t in transactions[i:i+batch_size]
                ]
                worksheet.append_rows(batch)

                # No sleep after last batch
                if i + batch_size < len(transactions):
                    time.sleep(1)

            total_income = sum(t['amount']
                               for t in transactions if t['type'] == 'income')
            total_expense = sum(t['amount']
                                for t in transactions if t['type'] == 'expense')
            savings = total_income - total_expense
            expense_rate = (
                total_expense / total_income) if total_income > 0 else 0
            savings_rate = (
                savings / total_income) if total_income > 0 else 0

            worksheet.format('B2:B4', {'numberFormat': {
                             'type': 'CURRENCY', 'pattern': '€#,##0.00'}, "textFormat": {'bold': True, 'fontSize': 12}})
            worksheet.format('C8:C31', {'numberFormat': {
                             'type': 'CURRENCY', 'pattern': '€#,##0.00'}})
            worksheet.format('A7:E7', {"textFormat": {
                'bold': True, 'fontSize': 12}})

            worksheet.update('A2:A4', [['Total Income:'], [
                             'Total Expenses:'], ['Savings:']])
            worksheet.update('B2:B4', [[total_income], [
                             total_expense], [savings]])
            worksheet.update('C2:C4', [[1], [
                             expense_rate], [savings_rate]])
            worksheet.format('C2:C4', {'numberFormat': {
                'type': 'PERCENT',
                'pattern': '0%'}})

            print(
                f"\nSuccessfully updated {len(transactions)} transactions in Google Sheets")
        else:
            print("\nNo transactions to update in Google Sheets")

    except Exception as e:
        print(f"\nError in Google Sheets operation: {str(e)}")


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
