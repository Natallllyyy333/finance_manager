# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
from collections import defaultdict
import time
from datetime import datetime
from itertools import zip_longest
from gspread_formatting import *
from gspread_formatting import (
    cellFormat,
    format_cell_range,
    Padding,
    set_column_width
)
from gspread.utils import rowcol_to_a1
import warnings
warnings.filterwarnings('ignore', category=DeprecationWarning)
import os
import sys
from io import StringIO
import json
from google.oauth2 import service_account
import threading
from flask import Flask, request, render_template_string


app = Flask(__name__)

DAILY_NORMS = {
        'Rent': 50.0,
        'Gym': 3.0,        # 45 / 30
        'Groceries': 3,    # 90 / 30
        'Transport': 0.27,  # 8 / 30
        'Entertainment': 0.17,  # 5 / 30
        'Utilities': 2.0,
        'Shopping': 3.33,  # 100 / 30
        'Dining': 10.00
    }

def sync_google_sheets_operation(month_name, table_data):
    """Синхронная версия Google Sheets операции"""
    try:
        print(f"📨 Starting Google Sheets sync for {month_name}")
        print(f"📊 Data to write: {len(table_data)} rows")
        
        time.sleep(2)
        
        # 1. Authentification
        print("🔑 Getting credentials...")
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials available")
            return False
            
        print("✅ Credentials obtained, authorizing...")
        gc = gspread.authorize(creds)
        
        print("✅ Authorized, opening spreadsheet...")
        # 2. Open target table by ID
        try:
            target_spreadsheet = gc.open_by_key('1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ')
            print("✅ Spreadsheet opened successfully")
        except Exception as e:
            print(f"❌ Error opening spreadsheet: {e}")
            return False
            
        try:
            summary_sheet = target_spreadsheet.worksheet('SUMMARY')
            print("✅ SUMMARY worksheet accessed")
        except Exception as e:
            print(f"❌ Error accessing SUMMARY worksheet: {e}")
            return False
        
        print("📋 Getting headers...")
        # 3. Get current headers
        headers = summary_sheet.row_values(2)
        print(f"📝 Current headers: {headers}")

        # 4. Normalizing month name for comparison
        normalized_month = month_name.capitalize()
        print(f"🔍 Looking for column: {normalized_month}")

        # 5. Find the month column
        month_col = None
        for i, header in enumerate(headers, 1):
            if header == normalized_month:
                month_col = i
                print(f"✅ Found existing column for {normalized_month} at position: {month_col}")
                break

        if month_col is None:
            print("🔍 No existing column found, looking for empty column...")
            # Find first empty column
            for i, header in enumerate(headers, 1):
                if not header.strip():  # Empty column
                    month_col = i
                    print(f"✅ Found empty column at position: {month_col}")
                    print(f"📝 Creating new column for {normalized_month}...")
                    summary_sheet.update_cell(2, month_col, normalized_month)
                    summary_sheet.update_cell(3, month_col + 1, f"{normalized_month} %")
                    print(f"✅ Created new column for {normalized_month} at position: {month_col}")
                    break
        
        if month_col is None:
            print("🔍 No empty columns, adding at the end...")
            # Add new columns at the end
            month_col = len(headers) + 1
            if month_col > 37:
                print("❌ Column limit reached (37)")
                return False
            print(f"📝 Adding new column at position: {month_col}")
            summary_sheet.update_cell(2, month_col, normalized_month)
            summary_sheet.update_cell(3, month_col + 1, f"{normalized_month} %")
            print(f"✅ Added new column for {normalized_month} at position: {month_col}")
        
        print("📝 Preparing data for writing...")
        # 6. Prepare data to be written
        update_data = []
        for i, row_data in enumerate(table_data, start=4):
            if len(row_data) == 3:
                category, amount, percentage = row_data
                update_data.append({
                    'range': f"{gspread.utils.rowcol_to_a1(i, month_col)}",
                    'values': [[amount]]
                })
                update_data.append({
                    'range': f"{gspread.utils.rowcol_to_a1(i, month_col + 1)}", 
                    'values': [[percentage]]
                })
        
        print(f"📤 Ready to write {len(update_data)} cells")
        
        # 7. batch-query
        if update_data:
            print("⏳ Writing data to Google Sheets...")
            batch_size = 5
            for i in range(0, len(update_data), batch_size):
                batch = update_data[i:i+batch_size]
                summary_sheet.batch_update(batch)
                print(f"✅ Batch {i//batch_size + 1} written")
                if i + batch_size < len(update_data):
                    time.sleep(10)
            
            print("✅ All data written successfully!")
            
            # Format percentage column
            try:
                print("🎨 Formatting percentage column...")
                percent_col = month_col + 1
                start_row = 4
                end_row = start_row + len(table_data) - 1
                for row in range(start_row, end_row + 1):
                    cell_address = f"{rowcol_to_a1(row, percent_col)}"
                    summary_sheet.format(cell_address, {
                        "numberFormat": {"type": "PERCENT", "pattern": "0.00%"},
                        "horizontalAlignment": "CENTER"
                    })
                print("✅ Percentage column formatted")
            except Exception as format_error:
                print(f"⚠️ Formatting error: {format_error}")
        
        print("✅ Google Sheets update completed successfully!")
        return True
        
    except Exception as e:
        print(f"❌ Error in sync_google_sheets_operation: {e}")
        import traceback
        print(f"🔍 Traceback: {traceback.format_exc()}")
        return False
def async_google_sheets_operation(month_name, table_data):
    """Асинхронно выполняет Google Sheets операции"""
    try:
        print(f"🚀 Starting async Google Sheets operation for {month_name}")
        # Небольшая задержка для инициализации
        time.sleep(10)

        # if len(table_data) > 50:
        #     print("⚠️ Large dataset, using optimized approach")
        #     # Упрощенная логика для больших данных
        
        # Вызываем синхронную версию
        success = sync_google_sheets_operation(month_name, table_data)
        
        if success:
            print(f"✓ Асинхронная запись в Google Sheets завершена для {month_name}")
        else:
            print(f"✗ Ошибка при асинхронной записи в Google Sheets")
            
    except Exception as e:
        print(f"Async Google Sheets error: {e}")
        import traceback
        print(f"🔥 Traceback: {traceback.format_exc()}")

def get_google_credentials():
    """Get Google credentials from environment variables or file"""
    if "DYNO" in os.environ:
        print("🔑 Using environment credentials from Heroku")
        # import json
        service_account_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
        if service_account_json:
            print("✅ GOOGLE_SERVICE_ACCOUNT_JSON found")
            try:
                creds_dict = json.loads(service_account_json)
                from google.oauth2 import service_account
                SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
                         'https://www.googleapis.com/auth/drive']
                return service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
            except json.JSONDecodeError:
                raise Exception("Invalid JSON in GOOGLE_SERVICE_ACCOUNT_JSON")
        else:
            print("❌ GOOGLE_SERVICE_ACCOUNT_JSON not found")
            
            raise Exception("GOOGLE_SERVICE_ACCOUNT_JSON environment variable not found")
            
    else:
        # Locally from file
        print("🔑 Using local credentials file")
        from google.oauth2 import service_account
        return service_account.Credentials.from_service_account_file('creds.json')
    


def load_transactions(filename):
        """Load and categorize transactions with daily tracking"""
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
            print(f"Error: File '{filename}' not found")
            exit()
        return transactions, daily_categories


def categorize(description):
    """Categorize transaction based on description."""
    desc = description.lower()
    categories = {
        'Salary': ['salary','wages'],
        'Bonus': ['bonus', 'tip', 'reward'],
        'Other income': ['stocks', 'exchange', 'earnings', 'prize'],
        'Rent': ['rent', 'monthly rent'],
        'Groceries': ['supermarket', 'grocery', 'food'],
        'Dining': ['restaurant', 'cafe', 'coffee'],
        'Transport': ['bus', 'train', 'taxi', 'uber'],
        'Entertainment': ['movie', 'netflix', 'concert'],
        'Utilities': ['electricity', 'water', 'gas', 'internet', 'phone'],
        'Gym': ['gym', 'Gym Membership' 'fitness', 'yoga'],
        'Shopping': ['clothing', 'electronics', 'shopping', 'Supermarket'],
        'Health': ['pharmacy', 'doctor', 'health', 'dentist'],
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

def get_month_column_name(month_input):
    """Привести название месяца к стандартному формату"""
    month = month_input.strip().capitalize()
    month_mapping = {
        'Jan': 'January', 'Feb': 'February', 'Mar': 'March',
        'Apr': 'April', 'May': 'May', 'Jun': 'June',
        'Jul': 'July', 'Aug': 'August', 'Sep': 'September',
        'Oct': 'October', 'Nov': 'November', 'Dec': 'December'
    }
    return month_mapping.get(month, month)
def analyze(transactions, daily_categories, month):
    """Perform financial analysis with daily tracking"""

    analysis = {
        'income': 0, 'expenses': 0, 'categories': defaultdict(float),
        'income_categories': defaultdict(float),
        'month': month, 'daily_categories': daily_categories,
        'days_count':  30,
        'daily_averages': defaultdict(float),
        'norms_violations': []
    }
    for t in transactions:
        if t['type'] == 'income':
            analysis['income'] += t['amount']
            analysis['income_categories'][t['category']] += t['amount']
        else:
            analysis['expenses'] += t['amount']
            analysis['categories'][t['category']] += t['amount']
    # Calculate daily averages
    for category, total in analysis['categories'].items():
        daily_avg = total / analysis['days_count']
        analysis['daily_averages'][category] = daily_avg
        if category in DAILY_NORMS:
            if daily_avg > DAILY_NORMS[category] * 1.1:  # 10% over norm
                analysis['norms_violations'].append(
                    f"Daily average for {category}"
                    f" overspent: {daily_avg:.2f}€ "
                    f"vs norm: {DAILY_NORMS[category]:.2f}€"
                )
    analysis['savings'] = analysis['income'] - analysis['expenses']
    return analysis


def terminal_visualization(data):
    """Visualize financial data in terminal."""
    # Header
    print(
        f" {data['month'].upper()} FINANCIAL OVERVIEW ".center(77, "="))
    # Summary bars
    expense_rate = (data['expenses'] / data['income']
                    * 100) if data['income'] > 0 else 0
    savings_rate = (data['savings'] / data['income']
                    * 100) if data['income'] > 0 else 0
    income_bar = "■" * int(data['income'] / max(data['income'], 1) * 20)
    print(f"Income: {data['income']:8.2f}€ [{income_bar}] 100%")
    expense_bar = "■" * int(data['expenses'] / max(data['income'], 1) * 20)
    print(f"Expenses: {data['expenses']:8.2f}€ ["
        f"{expense_bar}] {expense_rate:.1f}%")
    savings_bar = "■" * int(data['savings'] / max(data['income'], 1) * 20)
    print(f"Savings: {data['savings']:8.2f}€ ["
        f"{savings_bar}] {savings_rate:.1f}%")
    # Categories breakdown
    print(f" EXPENSE CATEGORIES ".center(77, '-'))
    top_categories = sorted(data['categories'].items(),
                            key=lambda x: x[1], reverse=True)[:9]  # 9 for 3 columns
    # Split into three columns
    col1 = top_categories[0:3]
    col2 = top_categories[3:6]
    col3 = top_categories[6:9]

    # Fixed width for each column component
    NAME_WIDTH = 10    # Category name
    AMOUNT_WIDTH = 9   # Amount (6.2f + € + space)
    BAR_WIDTH = 6      # Bar visualization
    
    # Total column width including spacing
    COLUMN_WIDTH = NAME_WIDTH + 1 + AMOUNT_WIDTH + 1 + BAR_WIDTH  # +2 for spaces

# Display three columns
    for (cat1, amt1), (cat2, amt2), (cat3, amt3) in zip_longest(col1, col2, col3, fillvalue=(None, 0)):
        line = ""
        if cat1:
            pct1 = (amt1 / data['expenses'] * 100) if data['expenses'] > 0 else 0
            bar1 = "■" * min(int(pct1 / 1), BAR_WIDTH)

            col1_text = f"{cat1[:NAME_WIDTH]:<{NAME_WIDTH}} {amt1:6.2f}€ {bar1:<{BAR_WIDTH}}"
            line += col1_text.ljust(COLUMN_WIDTH)
            
        else:
            line += " " *  COLUMN_WIDTH
        line += ""
                
        if cat2:
            
            pct2 = (amt2 / data['expenses'] * 100) if data['expenses'] > 0 else 0
            bar2 = "■" * min(int(pct2 / 1), BAR_WIDTH)
            col2_text = f"{cat2[:NAME_WIDTH]:<{NAME_WIDTH}} {amt2:6.2f}€ {bar2:<{BAR_WIDTH}}"
            line += col2_text.ljust(COLUMN_WIDTH)
            
        else:
            line += " " * COLUMN_WIDTH 
        line += ""
        

        if cat3:
            
            pct3 = (amt3 / data['expenses'] * 100) if data['expenses'] > 0 else 0
            bar3 = "■" * min(int(pct3 / 1), BAR_WIDTH)
            col3_text = f"{cat3[:NAME_WIDTH]:<{NAME_WIDTH}} {amt3:6.2f}€ {bar3:<{BAR_WIDTH}}"
            line += col3_text.ljust(COLUMN_WIDTH)
        
        print(line)
    print(f" DAILY SPENDING and NORMS ".center(77, '='))
    sorted_categories = sorted(
        [
            (cat, avg)
            for cat, avg in data['daily_averages'].items()
            if cat in DAILY_NORMS
        ],
        key=lambda x: x[1] - DAILY_NORMS.get(x[0], 0),
        reverse=True
    )[:3]
    for category, avg in sorted_categories:
        norm = DAILY_NORMS.get(category, 0)
        diff = avg - norm
        print(f"{category:<12} Avg: {avg:5.2f}€  Norm: {norm: 5.2f}€ "
            f"{'▲' if diff > 0 else '▼'} {abs(diff):.2f}€ "
            )


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
            # Add top 3 norms violations
            recs.extend(data['norms_violations'][:3])
        # Ensure minimum recommendations
        if len(recs) < 3:
            recs.extend([
                "Plan meals weekly to reduce grocery costs",
                "Use public transport more frequently",

            ])
        return recs[:3]  # Return only top 5 recommendations


def prepare_summary_data(data, transactions):
    """Prepare the data for the SUMMARY section - all categories and totals."""
    #
    all_categories = [
        'TOTAL INCOME',
        'TOTAL EXPENSES',
        'SAVINGS',
        '',
        'INCOME CATEGORIES:',
        'Salary',
        'Bonus',
        'Other Income',
        '',
        'EXPENSE CATEGORIES:',
        'Rent',
        'Groceries',
        'Dining',
        'Transport',
        'Entertainment',
        'Utilities',
        'Gym',
        'Shopping',
        'Health',
        'Insurance',
        'Education',
        'Travel',
        'Car',
        'Other'
    ]
    # Collecting data by income
    income_by_category = defaultdict(float)
    for t in transactions:
        if t['type'] == 'income':
            income_by_category[t['category']] += t['amount']

    # Colleciting data by expense
    expenses_by_category = defaultdict(float)
    for t in transactions:
        if t['type'] == 'expense':
            expenses_by_category[t['category']] += t['amount']

    # Preparing totals
    table_data = []
    for category in all_categories:
        if category == 'TOTAL INCOME':
            table_data.append([category, data['income'], 1.0])
        elif category == 'TOTAL EXPENSES':
            percentage = (data['expenses'] / data['income']
                        if data['income'] > 0 else 0)
            table_data.append([category, data['expenses'], percentage])

        elif category == 'SAVINGS':
            percentage = (data['savings'] / data['income']
                        if data['income'] > 0 else 0)
            table_data.append([category, data['savings'], percentage])

        elif category in ['', 'INCOME CATEGORIES:', 'EXPENSE CATEGORIES:']:

            table_data.append([category, '', ''])
        elif category in income_by_category:
            amount = income_by_category[category]
            percentage = (amount / data['income']
                        if data['income'] > 0 else 0)
            table_data.append([category, amount, percentage])
        elif category == 'Salary':
            matched = False
            for income_cat in income_by_category:
                amount = income_by_category[income_cat]
                percentage = (amount / data['income']
                            if data['income'] > 0 else 0)
                table_data.append([category, amount, percentage])
                matched = True
                break
            if not matched:
                if category == 'Salary':
                    for income_cat in income_by_category:
                        if ('salary' in income_cat.lower() or
                                'income' in income_cat.lower()):
                            amount = income_by_category[income_cat]
                            percentage = (amount / data['income']
                                        if data['income'] > 0 else 0)
                        table_data.append([category, amount, percentage])
                        matched = True
                        break
            if not matched:
                table_data.append([category, 0, 0])
        elif category in expenses_by_category:
            amount = expenses_by_category[category]
            percentage = amount / \
                data['expenses'] if data['expenses'] > 0 else 0
            table_data.append([category, amount, percentage])
        else:
            table_data.append([category, 0, 0])
    return table_data





def write_to_target_sheet(table_data, month_name):
    """Записать данные в целевую таблицу SUMMARY"""
    try:
        if not table_data:
            print("✗ No data to write to target sheet")
            return False
        
        # Проверяем размер данных
        if len(table_data) > 50:
            print(f"⚠️ Large dataset ({len(table_data)} rows), simplifying update")
            # Упрощаем данные для больших наборов
            simplified_data = []
            for row in table_data:
                if row[0] in ['TOTAL INCOME', 'TOTAL EXPENSES', 'SAVINGS']:
                    simplified_data.append(row)
                elif row[0] and not any(x in row[0] for x in ['CATEGORIES', '']):
                    simplified_data.append([row[0], row[1], 0])
            table_data = simplified_data

        # Если в Heroku, запускаем асинхронно
        if "DYNO" in os.environ:
            thread = threading.Thread(target=async_google_sheets_operation, args=(month_name, table_data))
            thread.daemon = True
            thread.start()
            print("Google Sheets operation started in background")
            return True
        else:
            # Локально выполняем синхронно
            return sync_google_sheets_operation(month_name, table_data)
            
    except Exception as e:
        print(f"✗ Ошибка записи в SUMMARY: {e}")
        return False
    # try:
    #     # 1. Authentification
    #     creds = get_google_credentials()
    #     gc = gspread.authorize(creds)

    #     # 2. Open target table by ID
    #     target_spreadsheet = gc.open_by_key(
    #         '1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ')
    #     summary_sheet = target_spreadsheet.worksheet('SUMMARY')

    #     # 3. Get current headers
    #     headers = summary_sheet.row_values(2)

    #     # 4. Normalizing month name for comparison
    #     normalized_month = month_name.capitalize()

    #     # 4. Find the month column
    #     month_col = None
    #     for i, header in enumerate(headers, 1):  # Начинаем с 1 столбца
    #         if header == normalized_month:
    #             month_col = i
    #             break

    #     if month_col is None:
    #         # Find first empty column
    #         for i, header in enumerate(headers, 1):
    #             if not header.strip():  # Empty column
    #                 month_col = i
    #                 """Write the name of the month
    #                 into the second row in the cell number month_col"""
    #                 summary_sheet.update_cell(2, month_col, normalized_month)
    #                 summary_sheet.update_cell(
    #                     3, month_col + 1, f"{normalized_month} %")
    #                 print(
    #                     f"Создан новый столбец для "
    #                     f"{normalized_month} в позиции: {month_col}")
    #                 break
    #     if month_col is None:
    #         # Add new columns at the end
    #         month_col = len(headers) + 1
    #         if month_col > 37:  # Проверка ограничения Google Sheets
    #             print("✗ Достигнут лимит столбцов (37)")
    #             return False
    #         summary_sheet.update_cell(2, month_col, normalized_month)
    #         summary_sheet.update_cell(
    #             3, month_col + 1, f"{normalized_month} %")
    #         time.sleep(20)
    #     # 5. Prepare data to be written
    #     update_data = []
    #     num_rows = len(table_data)
    #     for i, row_data in enumerate(table_data, start=4):
    #         if len(row_data) == 3:
    #             category, amount, percentage = row_data
    #         elif len(row_data) == 2:
    #             category, amount = row_data
    #             percentage = 0
    #         else:
    #             continue
    #         update_data.append({
    #             'range': f"{gspread.utils.rowcol_to_a1(i, month_col)}",
    #             'values': [[amount]]
    #         })
    #         update_data.append({
    #             'range': f"{gspread.utils.rowcol_to_a1(i, month_col + 1)}",
    #             'values': [[percentage]]
    #         })

    #         # 6. batch-query
    #     if update_data:
    #         batch_size = 10
    #         for i in range(0, len(update_data), batch_size):
    #             batch = update_data[i:i+batch_size]
    #             summary_sheet.batch_update(batch)
    #             if i + batch_size < len(update_data):
    #                 time.sleep(20)
    #         time.sleep(20)
    #         try:
    #             percent_col = month_col + 1
    #             start_row = 4
    #             end_row = start_row + len(table_data) - 1
    #             for row in range(start_row, end_row + 1):
    #                 cell_address = f"{rowcol_to_a1(row, percent_col)}"
    #                 summary_sheet.format(cell_address, {
    #                     "numberFormat": {
    #                         "type": "PERCENT",
    #                         "pattern": "0.00%"
    #                     },
    #                     "horizontalAlignment": "CENTER"
    #                 })
    #                 time.sleep(0.1)
    #         except Exception as format_error:
    #             print(f"⚠️ Percent column formating error: {format_error}")
    #         time.sleep(20)
    #     return True
    # except Exception as e:
    #     print(f"✗ Ошибка записи в SUMMARY: {e}")
    #     return False


def main():
    
    print(f" PERSONAL FINANCE ANALYZER ".center(77, "="))
    MONTH = input(
        "Enter the month (e.g. 'March, April, May'): ").strip().lower()
    FILE = f"hsbc_{MONTH}.csv"
    print(f"Loading file: {FILE}")

    transactions, daily_categories = load_transactions(FILE)
    if not transactions:
        print(f"No transactions found")
        return
    data = analyze(transactions, daily_categories, MONTH)
    terminal_visualization(data)
    # Recommendations
    print(f" DAILY SPENDING RECOMMENDATIONS ".center(77, '='))
    for i, rec in enumerate(generate_daily_recommendations(data), 1):
        print(f"{i}. {rec}")
    # Optional Google Sheets update
    try:
        # Authenticate and open Google Sheets
        creds = get_google_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")
        # Check if worksheet exists
        worksheet = None
        try:
            worksheet = sh.worksheet(MONTH)
            print(f"\n"+ f"Worksheet '{MONTH}' found. Updating...")
        except gspread.WorksheetNotFound:
            print(f"Worksheet for {MONTH} not found. Creating a new one...")
            # First check if we've reached the sheet limit (max 200 sheets)
            if len(sh.worksheets()) >= 200:
                raise Exception("Maximum number of sheets (200) reached")
            """Check if sheet exists but with
            different case (e.g. "march" vs "March")"""
            existing_sheets = [ws.title for ws in sh.worksheets()]
            if MONTH.lower() in [sheet.lower() for sheet in existing_sheets]:
                # Find the existing sheet with case-insensitive match
                for sheet in sh.worksheets():
                    if sheet.title.lower() == MONTH.lower():
                        worksheet = sheet
                        print(
                            f"Using existing worksheet '{sheet.title}'"
                            f"(case difference)")
                        break
            else:
                # Create new worksheet with unique name if needed
                try:
                    worksheet = sh.add_worksheet(
                        title=MONTH, rows="100", cols="20")
                    print(f"New worksheet '{MONTH}' created successfully.")
                except gspread.exceptions.APIError as e:
                    if "already exists" in str(e):
                        """If we get here, it means the sheet
                        exists but wasn't found earlier"""
                        worksheet = sh.worksheet(MONTH)
                        print(f"Worksheet '{MONTH}' exists. Using it.")
                    else:
                        raise e
        if worksheet is None:
            raise Exception("Failed to access or create worksheet")
        # Clear existing data (keep headers)
        all_values = worksheet.get_all_values()
        if len(all_values) > 1:
            worksheet.delete_rows(1, len(all_values)+1)
        time.sleep(20)
        all_data = [["Date", "Description", "Amount", "Type", "Category"]]
        for t in transactions:
            all_data.append([t['date'], t['desc'][:50],
                            t['amount'], t['type'], t['category']])
        worksheet.update('A7', all_data)
        time.sleep(3)
        total_income = sum(t['amount']
                        for t in transactions if t['type'] == 'income')
        total_expense = sum(t['amount']
                            for t in transactions if t['type'] == 'expense')
        savings = total_income - total_expense
        expense_rate = (
            total_expense / total_income) if total_income > 0 else 0
        savings_rate = (
            savings / total_income) if total_income > 0 else 0
        # Группируем все операции форматирования
        try:
            format_operations = [
                ('B2:B4', {
                    'numberFormat': {'type': 'CURRENCY', 'pattern': '€#,##0.00'},
                    "textFormat": {'bold': True, 'fontSize': 12}
                }),
                ('C8:C31', {'numberFormat': {'type': 'CURRENCY', 'pattern': '€#,##0.00'}}),
                ('A7:E7', {
                    "textFormat": {'bold': True, 'fontSize': 12},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                }),
                ('C2:C4', {'numberFormat': {'type': 'PERCENT', 'pattern': '0%'}}),
                ('G7:I7', {
                    "textFormat": {"bold": True, "fontSize": 12},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                }),
                (f'H8:H{end_row}', {"numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}}),
                (f'I8:I{end_row}', {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}})
            ]
            
            for range_name, format_dict in format_operations:
                worksheet.format(range_name, format_dict)
                time.sleep(10)  # Уменьшите задержку между операциями
            
        except Exception as e:
            print(f"Formatting error: {e}")
    # Продолжаем выполнение даже если форматирование не удалось
        # worksheet.format('B2:B4', {
        #     'numberFormat': {
        #         'type': 'CURRENCY',
        #         'pattern': '€#,##0.00'},
        #     "textFormat": {
        #         'bold': True,
        #         'fontSize': 12
        #         }
        #         })
        # time.sleep(20)
        # worksheet.format('C8:C31', {'numberFormat': {
        #     'type': 'CURRENCY', 'pattern': '€#,##0.00'}})
        # time.sleep(20)
        # worksheet.format('A7:E7', {"textFormat": {
        #     'bold': True, 'fontSize': 12},
        #     "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}})
        # time.sleep(20)
        # worksheet.update('A2:A4', [['Total Income:'], [
        #     'Total Expenses:'], ['Savings:']])
        # time.sleep(20)
        # worksheet.update('B2:B4', [[total_income], [
        #     total_expense], [savings]])
        # time.sleep(20)
        # worksheet.update('C2:C4', [[1], [
        #     expense_rate], [savings_rate]])
        # time.sleep(20)
        # worksheet.format('C2:C4', {'numberFormat': {
        #     'type': 'PERCENT',
        #     'pattern': '0%'}})
        # time.sleep(20)
        if transactions:
            expenses_by_category = defaultdict(float)
            for t in transactions:
                if t['type'] == 'expense':
                    expenses_by_category[t['category']] += t['amount']
        sorted_categories = sorted(
            expenses_by_category.items(), key=lambda x: x[1], reverse=True)
        category_data = []
        total_expenses = data['expenses']
        for category, amount in sorted_categories:
            percentage = (amount / total_expenses *
                        100) if total_expenses > 0 else 0
            category_data.append([
                f"{category}: {amount:.2f}€ ({percentage:.1f}%)"])
            time.sleep(5)
        if category_data:
            last_row = 7 + len(category_data)
            last_row_transactions = 7 + len(transactions)
            table_data = []
            table_data = prepare_summary_data(data, transactions)
            time.sleep(20)
            if last_row < 7 + len(table_data):
                rows_to_add = (7 + len(table_data)) - last_row
                worksheet.add_rows(rows_to_add)
                time.sleep(20)
            MONTH_NORMALIZED = get_month_column_name(MONTH)
            write_to_target_sheet(table_data, MONTH_NORMALIZED)  # Убрали success =
            time.sleep(20)
            # MONTH_NORMALIZED = get_month_column_name(
            #     MONTH)
            # success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
            # time.sleep(20)
            category_headers = [['Category', 'Amount', 'Percentage']]
            worksheet.update('G7:I7', category_headers)
            time.sleep(20)
            category_table_data = []
            for category, amount, percentage in table_data:
                if isinstance(percentage, (int, float)):
                    category_table_data.append([category, amount, percentage])
                else:
                    category_table_data.append([category, amount, 0])
            end_row = 7 + len(category_table_data)
            current_rows = worksheet.row_count
            if end_row > current_rows:
                rows_to_add = end_row - current_rows
                worksheet.add_rows(rows_to_add)
                time.sleep(20)
            worksheet.update(f'G8:I{end_row}', category_table_data)
            time.sleep(20)
            category_end_row = 7 + len(table_data)
            if category_end_row > worksheet.row_count:
                rows_to_add = category_end_row - worksheet.row_count
                worksheet.add_rows(rows_to_add)
                time.sleep(20)
            worksheet.update(f'G8:I{category_end_row}', table_data)
            time.sleep(20)
            worksheet.format('G7:I7', {
                "textFormat": {"bold": True, "fontSize": 12},
                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
            })
            worksheet.format(
                f'H8:H{end_row}',
                {
                    "numberFormat": {
                        "type": "CURRENCY",
                        "pattern": "€#,##0.00"
                                    }
                })
            time.sleep(20)
            worksheet.format(f'I8:I{end_row}', {
                "numberFormat": {
                    "type": "PERCENT",
                    "pattern": "0.00%"
                }
            })
            time.sleep(20)
            column_formats = [
                (f'A8:A{last_row_transactions}', {"backgroundColor": {
                "red": 0.90, "green": 0.90, "blue": 0.90}}),
                (f'B8:B{last_row_transactions}', {"backgroundColor": {
                "red": 0.96, "green": 0.96, "blue": 0.96}}),
                (f'C8:C{last_row_transactions}', {"backgroundColor": {
                "red": 0.94, "green": 0.94, "blue": 0.94}}),
                (f'D8:D{last_row_transactions}', {"backgroundColor": {
                "red": 0.92, "green": 0.92, "blue": 0.92}}),
                (f'E8:E{last_row_transactions}', {"backgroundColor": {
                "red": 0.90, "green": 0.90, "blue": 0.90}})
            ]
            for range_, format_ in column_formats:
                worksheet.format(range_, format_)
            time.sleep(20)
            category_column_formats = [
                (f'G8:G{end_row}', {"backgroundColor": {
                "red": 0.94, "green": 0.94, "blue": 0.94}}),
                (f'H8:H{end_row}', {"backgroundColor": {
                "red": 0.96, "green": 0.96, "blue": 0.96}}),
                (f'I8:I{end_row}', {"backgroundColor": {
                "red": 0.94, "green": 0.94, "blue": 0.94}})
            ]
            for range_, format_ in category_column_formats:
                worksheet.format(range_, format_)
            time.sleep(20)
            worksheet.update('A2:A4', [['Total Income:'], [
                            'Total Expenses:'], ['Savings:']])
            time.sleep(20)
            worksheet.update('B2:B4', [[total_income], [
                            total_expense], [savings]])
            time.sleep(20)
            worksheet.update('C2:C4', [[1], [
                            expense_rate], [savings_rate]])
            time.sleep(20)
            worksheet.format('C2:C4', {'numberFormat': {
                'type': 'PERCENT',
                'pattern': '0%'}})
            
            # Упрощенное форматирование границ - делаем одним вызовом если возможно
            try:
                border_format = {
                    "borders": {
                        "top": {"style": "SOLID", "width": 1, "color": {"red": 0.6, "green": 0.6, "blue": 0.6}},
                        "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0.6, "green": 0.6, "blue": 0.6}},
                        "left": {"style": "SOLID", "width": 1, "color": {"red": 0.6, "green": 0.6, "blue": 0.6}},
                        "right": {"style": "SOLID", "width": 1, "color": {"red": 0.6, "green": 0.6, "blue": 0.6}}
                    }
                }
                
                # Форматируем все таблицы одним batch если возможно
                worksheet.format(f'A7:E{7 + len(transactions)}', border_format)
                worksheet.format(f'G7:I{end_row}', border_format)
                worksheet.format('A2:C4', border_format)
                time.sleep(15)
                
            except Exception as e:
                print(f"Border formatting error: {e}")
            # border_style = {
            #     "style": "SOLID",
            #     "width": 1,
            #     "color": {"red": 0.6, "green": 0.6, "blue": 0.6}
            # }
            # border_format = {
            #     "borders": {
            #         "top": border_style,
            #         "bottom": border_style,
            #         "left": border_style,
            #         "right": border_style
            #     }
            # }
            # tables = [
            #     f'A7:E{7 + len(transactions)}',  # Основная таблица транзакций
            #     f'G7:I{end_row}',  # Таблица категорий
            #     'A2:C4'                          # Блок с итогами
            # ]
            # for table_range in tables:
            #     worksheet.format(table_range, border_format)
            # time.sleep(20)
            header_bottom_border = {
                "borders": {
                    "bottom": {
                        "style": "SOLID",
                        "width": 2,
                        "color": {"red": 0.4, "green": 0.4, "blue": 0.4}
                    }
                }
            }
            header_left_border = {
                "borders": {
                    "left": {
                        "style": "SOLID",
                        "width": 2,
                        "color": {"red": 0.4, "green": 0.4, "blue": 0.4}
                    }
                }
            }
            header_top_border = {
                "borders": {
                    "top": {
                        "style": "SOLID",
                        "width": 2,
                        "color": {"red": 0.4, "green": 0.4, "blue": 0.4}
                    }
                }
            }
            worksheet.format('A7:E7', header_bottom_border)
            time.sleep(20)
            worksheet.format('G7:I7', header_bottom_border)
            time.sleep(20)
            worksheet.format('D2:D4', header_left_border)
            time.sleep(20)
            worksheet.format('A2:C4', border_format)
            time.sleep(20)
            worksheet.format('A1:C1', header_bottom_border)
            time.sleep(20)
            worksheet.format('A4:C4', border_format)
            time.sleep(20)
            worksheet.format('A5:C5', header_top_border)
            time.sleep(20)
            recommendations = generate_daily_recommendations(data)
            time.sleep(20)
            rec_headers = ["Priority", "Recommendation"]
            rec_data = [[f"{i+1}.", rec]
                        for i, rec in enumerate(recommendations)]
            rec_start_row = 7
            rec_start_col = 11
            worksheet.update(
                values=[rec_headers],
                range_name=f"K{rec_start_row}:L{rec_start_row}"
            )
            time.sleep(20)
            for i, row in enumerate(rec_data, start=rec_start_row+1):
                worksheet.update(f"K{i}:L{i}", [row])
            time.sleep(20)
            fmt = cellFormat(
                horizontalAlignment='CENTER',
                padding=Padding(top=8, right=12, bottom=8, left=12),
                wrapStrategy='WRAP'
            )
            time.sleep(20)
            worksheet.format(
                f"K{rec_start_row}:L{rec_start_row}",
                {
                    "textFormat": {"bold": True, "fontSize": 12},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    }
                }
            )
            time.sleep(20)
            worksheet.format(
                f"K{rec_start_row + 1}:L{rec_start_row + len(rec_data)}",
                {
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1}
                    },
                    "wrapStrategy": "WRAP",
                }
            )
            time.sleep(20)
            set_column_width(worksheet, 'A', 120)
            set_column_width(worksheet,  'C',  80)
            set_column_width(worksheet,  'D',  80)
            set_column_width(worksheet, 'E', 80)
            set_column_width(worksheet,  'K',  90)
            set_column_width(worksheet,   'G',  200)
            set_column_width(worksheet,  'H', 80)
            set_column_width(worksheet, 'B', 200)
            set_column_width(worksheet, 'L', 300)
            set_column_width(worksheet, 'F', 30)
            set_column_width(worksheet, 'J', 30)
            time.sleep(20)
            worksheet.update(f"K6", [['DAILY RECOMMENDATIONS']])
            time.sleep(20)
            worksheet.format("K6", {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            time.sleep(20)
            worksheet.merge_cells(f"K6:L6")
            time.sleep(20)
            worksheet.update(f"A6", [['FINANCIAL OVERVIEW']])
            time.sleep(20)
            worksheet.format("A6", {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            time.sleep(20)
            worksheet.merge_cells(f"A6:E6")
            time.sleep(20)
            worksheet.update(f"G6", [['TRANSACTION CATEGORIES']])
            time.sleep(20)
            worksheet.format("G6", {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            time.sleep(20)
            worksheet.merge_cells(f"G6:I6")
            time.sleep(20)
            worksheet.format("A1:Z100", {"horizontalAlignment": "CENTER"})
            time.sleep(20)

            # MONTH_NORMALIZED = get_month_column_name(MONTH)
            # write_to_target_sheet(table_data, MONTH_NORMALIZED)  # Убрали success = 
            # print(
            #     f"Google Sheets update initiated for {len(transactions)} transactions")

            # MONTH_NORMALIZED = get_month_column_name(MONTH)
            # success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
            # print(
            #     f"Successfully updated {len(transactions)} "
            #     f"transactions in Google Sheets")
            
        else:
            print("\nNo transactions to update in Google Sheets")
    except Exception as e:
        print(f"\nError in Google Sheets operation: {str(e)}")


if "DYNO" in os.environ:
    # Режим Heroku - запускаем как веб-приложение
    
    
    HTML = '''
    <!DOCTYPE html>
<html>
<head>
    <title>Finance Analyzer</title>
    <style>
        body { font-family: Arial; margin: 40px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        input, button { padding: 10px; margin: 10px 0; font-size: 16px; }
        pre { 
            background: #2d2d2d; 
            color: #f8f8f2; 
            padding: 20px; 
            border-radius: 5px; 
            overflow: auto;
            white-space: pre-wrap;
            font-family: 'Courier New', monospace;
        }
        .success { color: #4CAF50; }
        .loading { color: #FF9800; }
    </style>
</head>
<body>
    <div class="container">
        <h1>💰 Personal Finance Analyzer</h1>
        <form method="POST">
            <input type="text" name="month" placeholder="Enter month (e.g. March, April, May)" required>
            <button type="submit">Analyze</button>
        </form>
        
        {% if result %}
        <h2>📊 Results for {{ month }}:</h2>
        <pre>{{ result }}</pre>
        <p class="loading">⏳ Google Sheets update in progress... Check logs for details.</p>
        {% endif %}
    </div>
</body>
</html>
    '''
@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    month = None

    if request.method == 'POST':
        month = request.form['month'].strip()
        
        # Быстрый анализ для немедленного отображения
        try:
            # Быстрая обработка для отображения
            FILE = f"hsbc_{month}.csv"
            transactions, daily_categories = load_transactions(FILE)
            
            if transactions:
                data = analyze(transactions, daily_categories, month)
                
                # Форматируем результат для веб-отображения
                result = f"""
    === {month.upper()} FINANCIAL ANALYSIS ===
    Income: {data['income']:.2f}€
    Expenses: {data['expenses']:.2f}€
    Savings: {data['savings']:.2f}€
    Savings Rate: {(data['savings']/data['income']*100 if data['income'] > 0 else 0):.1f}%

    Top Expenses:
    """
                # Добавляем топ категорий
                top_categories = sorted(data['categories'].items(), key=lambda x: x[1], reverse=True)[:5]
                for category, amount in top_categories:
                    result += f"{category}: {amount:.2f}€\n"
                
                result += "\nGoogle Sheets update in progress..."
                
            else:
                result = f"No transactions found for {month}"
                
        except Exception as e:
            result = f"Error: {str(e)}"
        
        # Запускаем полную фоновую обработку
        thread = threading.Thread(target=run_full_analysis, args=(month,))
        thread.daemon = True
        thread.start()
    
    return render_template_string(HTML, result=result, month=month)
    # @app.route('/', methods=['GET', 'POST'])
    # def index():
    #     result = None
    #     month = None
        
    #     if request.method == 'POST':
    #         month = request.form['month'].strip()
    #         # result = run_analysis(month)
    #         result = f"Analysis for {month} started in background. Check logs for details."

    #          # Запускаем в фоне
    #         thread = threading.Thread(target=run_analysis, args=(month,))
    #         thread.daemon = True
    #         thread.start()
        
    #     return render_template_string(HTML, result=result, month=month)
    
    def run_analysis(month):
        """Запускает анализ с виртуальным вводом"""
        try:
            print(f"Starting background analysis for {month}")
            # Сохраняем оригинальные потоки
            old_stdin = sys.stdin
            old_stdout = sys.stdout
            
            # Создаем виртуальные потоки
            sys.stdin = StringIO(month + '\n')
            sys.stdout = output_capture = StringIO()
            
            # Запускаем основную логику
            # main()
            print(f"=== {month.upper()} FINANCIAL ANALYSIS INITIATED ===")
            print("Data processing started in background...")
            print("Google Sheets update will complete shortly")
            print("Check Heroku logs for detailed results")

             # Запускаем полный анализ в фоне
            thread = threading.Thread(target=run_full_analysis, args=(month,))
            thread.daemon = True
            thread.start()
                
            # Получаем вывод
            output = output_capture.getvalue()
            
            return output
            
        except Exception as e:
            return f"Error: {str(e)}"
        finally:
            # Восстанавливаем потоки
            sys.stdin = old_stdin
            sys.stdout = old_stdout

def write_to_month_sheet(month_name, transactions, data):
    """Запись данных в лист месяца в формате как на скриншоте"""
    try:
        print(f"📊 Writing to {month_name} worksheet...")
        
        # 1. Authentification
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials for month sheet")
            return False
            
        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")
        
        # 2. Get or create worksheet
        try:
            worksheet = sh.worksheet(month_name)
            print(f"✅ Worksheet '{month_name}' found")
        except gspread.WorksheetNotFound:
            print(f"📝 Creating new worksheet '{month_name}'...")
            worksheet = sh.add_worksheet(title=month_name, rows="100", cols="20")
            print(f"✅ Worksheet '{month_name}' created")
        
        # 3. Clear existing data
        worksheet.clear()
        time.sleep(2)
        
        # 4. Create the layout as shown in the screenshot
        # Financial Overview header
        worksheet.update('A6', [['FINANCIAL OVERVIEW']])
        worksheet.merge_cells('A6:E6')
        worksheet.format('A6', {
            "textFormat": {"bold": True, "fontSize": 14},
            "horizontalAlignment": "CENTER"
        })
        
        # Table headers
        headers = ["Date", "Description", "Amount", "Type", "Category"]
        worksheet.update('A7', [headers])
        worksheet.format('A7:E7', {
            "textFormat": {"bold": True, "fontSize": 12},
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            }
        })
        
        # Write transactions
        all_data = []
        for t in transactions:
            all_data.append([t['date'], t['desc'][:30], t['amount'], t['type'], t['category']])
        
        if all_data:
            worksheet.update('A8', all_data)
        
        # Format transaction table
        last_transaction_row = 7 + len(transactions)
        worksheet.format(f'A8:E{last_transaction_row}', {
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            }
        })
        
        # Format currency columns
        worksheet.format(f'C8:C{last_transaction_row}', {
            "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
        })
        
        # Transaction Categories header
        worksheet.update('G6', [['TRANSACTION CATEGORIES']])
        worksheet.merge_cells('G6:I6')
        worksheet.format('G6', {
            "textFormat": {"bold": True, "fontSize": 14},
            "horizontalAlignment": "CENTER"
        })
        
        # Categories table headers
        category_headers = ["Category", "Amount", "Percentage"]
        worksheet.update('G7', [category_headers])
        worksheet.format('G7:I7', {
            "textFormat": {"bold": True, "fontSize": 12},
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            }
        })
        
        # Prepare and write category data
        table_data = prepare_summary_data(data, transactions)
        category_data = []
        for row in table_data:
            if row[0] and row[0] not in ['', 'INCOME CATEGORIES:', 'EXPENSE CATEGORIES:']:
                category_data.append([row[0], row[1], row[2]])
        
        if category_data:
            worksheet.update('G8', category_data)
        
        # Format category table
        last_category_row = 7 + len(category_data)
        worksheet.format(f'G8:I{last_category_row}', {
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            }
        })
        
        # Format currency and percentage columns
        worksheet.format(f'H8:H{last_category_row}', {
            "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
        })
        worksheet.format(f'I8:I{last_category_row}', {
            "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
        })
        
        # Daily Recommendations header
        worksheet.update('K6', [['DAILY RECOMMENDATIONS']])
        worksheet.merge_cells('K6:L6')
        worksheet.format('K6', {
            "textFormat": {"bold": True, "fontSize": 14},
            "horizontalAlignment": "CENTER"
        })
        
        # Recommendations headers
        rec_headers = ["Priority", "Recommendation"]
        worksheet.update('K7', [rec_headers])
        worksheet.format('K7:L7', {
            "textFormat": {"bold": True, "fontSize": 12},
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            }
        })
        
        # Write recommendations
        recommendations = generate_daily_recommendations(data)
        rec_data = []
        for i, rec in enumerate(recommendations, 1):
            rec_data.append([f"{i}", rec])
        
        if rec_data:
            worksheet.update('K8', rec_data)
        
        # Format recommendations table
        last_rec_row = 7 + len(rec_data)
        worksheet.format(f'K8:L{last_rec_row}', {
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            },
            "wrapStrategy": "WRAP"
        })
         # Calculate percentages
        expense_percentage = (data['expenses'] / data['income']) if data['income'] > 0 else 0
        savings_percentage = (data['savings'] / data['income']) if data['income'] > 0 else 0
        # Summary section at the top
        summary_data = [
            ["Total Income:", data['income'], 1.0],
            ["Total Expenses:", data['expenses'], expense_percentage],
            ["Savings:", data['savings'], savings_percentage]
        ]
        worksheet.update('A2', summary_data)
        
        # Format summary section
        worksheet.format('A2:B4', {
            "textFormat": {"bold": True},
            "borders": {
                "top": {"style": "SOLID", "width": 1},
                "bottom": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1}
            }
        })
        
        worksheet.format('B2:B4', {
            "numberFormat": {"type": "CURRENCY", "pattern": "€#,##0.00"}
        })
        
        # Set column widths to match the screenshot
        set_column_width(worksheet, 'A', 100)  # Date
        set_column_width(worksheet, 'B', 200)  # Description
        set_column_width(worksheet, 'C', 80)   # Amount
        set_column_width(worksheet, 'D', 80)   # Type
        set_column_width(worksheet, 'E', 100)  # Category
        set_column_width(worksheet, 'G', 150)  # Category name
        set_column_width(worksheet, 'H', 80)   # Amount
        set_column_width(worksheet, 'I', 80)   # Percentage
        set_column_width(worksheet, 'K', 60)   # Priority
        set_column_width(worksheet, 'L', 300)  # Recommendation
        
        print(f"✅ Successfully formatted {month_name} worksheet to match screenshot")
        return True
        
    except Exception as e:
        print(f"❌ Error writing to {month_name} worksheet: {e}")
        import traceback
        print(f"🔍 Traceback: {traceback.format_exc()}")
        return False
    
def run_full_analysis(month):
    """Полная обработка в фоновом режиме"""
    try:
        import sys
        sys.stdout = sys.__stdout__
        
        print(f"🚀 Starting FULL background analysis for {month}")
        # Восстанавливаем настоящий stdout для логирования
        old_stdout = sys.stdout
        sys.stdout = sys.__stdout__
        
        print(f"Starting FULL background analysis for {month}")
        
        # Перенесите сюда всю логику из вашей main() функции
        FILE = f"hsbc_{month}.csv"
        print(f"Loading file: {FILE}")

        transactions, daily_categories = load_transactions(FILE)
        if not transactions:
            print("No transactions found")
            return
            
        data = analyze(transactions, daily_categories, month)
        
        # Выводим основные результаты
        print(f"{month.upper()} ANALYSIS COMPLETED")
        print(f"Income: {data['income']:.2f}€")
        print(f"Expenses: {data['expenses']:.2f}€")
        print(f"Savings: {data['savings']:.2f}€")
        # 1. ЗАПИСЬ В ЛИСТ МЕСЯЦА
        print(f"📝 Writing to {month} worksheet...")
         # 1. ЗАПИСЬ В ЛИСТ МЕСЯЦА (новый формат)
        print(f"📝 Writing to {month} worksheet in screenshot format...")
        write_to_month_sheet(month, transactions, data)
        
        time.sleep(10)
        print("⏳ Starting Google Sheets update...")
        # Запускаем Google Sheets
        table_data = prepare_summary_data(data, transactions)
        MONTH_NORMALIZED = get_month_column_name(month)
        write_to_target_sheet(table_data, MONTH_NORMALIZED)
        print("🎉 All background tasks completed!")
        
    except Exception as e:
        print(f"Background analysis error: {e}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
    finally:
        if 'old_stdout' in locals():
            sys.stdout = old_stdout  

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
else:
    if __name__ == "__main__":
        main()
