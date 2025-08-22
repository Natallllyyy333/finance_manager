# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
import gspread
from collections import defaultdict
import time
from datetime import datetime
from itertools import zip_longest
from gspread_formatting import *
from gspread_formatting import cellFormat, format_cell_range, Padding, set_column_width


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

print("\n" + " PERSONAL FINANCE ANALYZER ".center(80, "="))
MONTH = input(
    "Enter the month ('April'): ").strip().lower()
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
        'income_categories': defaultdict(float),
        'month': MONTH, 'daily_categories': daily_categories,
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
                    f"Daily average for {category} overspent: {daily_avg:.2f}€ vs norm: {DAILY_NORMS[category]:.2f}€"
                )

    analysis['savings'] = analysis['income'] - analysis['expenses']
    return analysis


def terminal_visualization(data):
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
          int(data['income'] / max(data['income'], 1) * 20) + "]" + " " + "100%")
    print(f"\nExpenses: {data['expenses']:9.2f}€ [" + "■" *
          int(data['expenses'] / max(data['income'], 1) * 20) + "]" + " " + f"{expense_rate:.1f}%")
    print(f"Savings: {data['savings']:10.2f}€ [" + "■" *
          int(data['savings'] / max(data['income'], 1) * 20) + "]" + " " + f"{savings_rate:.1f}%")

    # Categories breakdown
    print("\n" + f" EXPENSE CATEGORIES ".center(80, '-'))
    # for cat, amount in sorted(data['categories'].items(), key=lambda x: x[1], reverse=True)[:10]:
    #     pct = amount / data['expenses'] * 100 if data['expenses'] > 0 else 0
    top_categories = sorted(data['categories'].items(),
                            key=lambda x: x[1], reverse=True)[:10]
    left_col = top_categories[:5]
    right_col = top_categories[5:]
    # Выводим две колонки с гистограммами (5 строк)
    for (left_cat, left_amt), (right_cat, right_amt) in zip_longest(
            left_col, right_col, fillvalue=(None, 0)):
        left_line = ""
        if left_cat:
            left_pct = left_amt/data['expenses'] * \
                100 if data['expenses'] > 0 else 0
            left_bar = "■" * int(left_pct/5)  # 1 символ ■ на 5%
            left_line = f"{left_cat[:10]:<10} {left_amt:6.2f}€ {left_bar}"

            right_line = ""
        if right_cat:
            right_pct = right_amt/data['expenses'] * \
                100 if data['expenses'] > 0 else 0
            right_bar = "■" * int(right_pct/5)
            right_line = f"{right_cat[:10]:<10} {right_amt:6.2f}€ {right_bar}"
            print(f"{left_line:<38}  {right_line}")

    print("\n" + f" DAILY SPENDING and NORMS ".center(80, '='))

    sorted_categories = sorted([(cat, avg) for cat, avg in data['daily_averages'].items(
    ) if cat in DAILY_NORMS], key=lambda x: x[1] - DAILY_NORMS.get(x[0], 0), reverse=True)[:5]

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
        return recs[:5]  # Return only top 5 recommendations


def prepare_summary_data(data, transactions):
    """Подготовить данные для записи в SUMMARY - все категории + итоги"""
    # Определяем ВСЕ возможные категории которые должны быть в SUMMARY
    all_categories = [
        'TOTAL INCOME',
        'TOTAL EXPENSES',
        'SAVINGS',
        '',  # пустая строка разделитель
        'INCOME CATEGORIES:',
        'Salary',
        'Bonus',
        'Other Income',
        '',  # пустая строка разделитель
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

    # Собираем данные по доходам
    income_by_category = defaultdict(float)
    for t in transactions:
        if t['type'] == 'income':
            income_by_category[t['category']] += t['amount']

    # Собираем данные по расходам
    expenses_by_category = defaultdict(float)
    for t in transactions:
        if t['type'] == 'expense':
            expenses_by_category[t['category']] += t['amount']

    # Подготавливаем итоговые данные
    table_data = []
    for category in all_categories:
        if category == 'TOTAL INCOME':
            table_data.append([category, data['income'], 1.0])
        elif category == 'TOTAL EXPENSES':
            percentage = data['expenses'] / data['income'] if data['income'] > 0 else 0
            table_data.append([category, data['expenses'], percentage])
            # table_data.append([category, data['expenses'], 0])
        elif category == 'SAVINGS':
            percentage = data['savings'] / data['income'] if data['income'] > 0 else 0
            table_data.append([category, data['savings'], percentage])
            # table_data.append([category, data['savings'], 0])
        elif category in ['', 'INCOME CATEGORIES:', 'EXPENSE CATEGORIES:']:
            # table_data.append([category, 0, 0])  # заголовки и разделители
            table_data.append([category, '', ''])  # заголовки и разделители
        elif category in income_by_category:
            amount = income_by_category[category]
            percentage = amount / data['income'] if data['income'] > 0 else 0
            table_data.append([category, amount, percentage])
        elif category == 'Salary':
            matched = False
            for income_cat in income_by_category:
                amount = income_by_category[income_cat]
                percentage = amount / data['income'] if data['income'] > 0 else 0
                table_data.append([category, amount, percentage])
                matched = True
                break
            if not matched:
                if category == 'Salary':
                    for income_cat in income_by_category:
                        if 'salary' in income_cat.lower() or 'income' in income_cat.lower():
                             amount = income_by_category[income_cat]
                             percentage = amount / data['income'] if data['income'] > 0 else 0
                             table_data.append([category, amount, percentage])
                             matched = True
                             break
            if not matched:
                table_data.append([category, 0, 0])
        

                             







        

        elif category in expenses_by_category:
            amount = expenses_by_category[category]
            percentage = amount / data['expenses'] if data['expenses'] > 0 else 0
            table_data.append([category, amount, percentage])
        
            # table_data.append([category, amount, 0])
        else:
            # table_data.append([category, 0, 0])  # для категорий без операций
            table_data.append([category, 0, 0])  # для категорий без операций
    return table_data


def get_month_column_name(month_input):
    """Привести название месяца к стандартному формату"""
    month = month_input.strip().capitalize()
    month_mapping = {
        'Jan': 'January', 'Feb': 'February', 'Mar': 'March',
        'Apr': 'April', 'May': 'May', 'Jun': 'June',
        'Jul': 'July', 'Aug': 'August', 'Sep': 'September',
        'Oct': 'October', 'Nov': 'November', 'Dec': 'December'
    }
    # Если короткое название - преобразовать в полное
    return month_mapping.get(month, month)
    # if month in month_mapping:
    #     return month_mapping[month]
    # return month


def write_to_target_sheet(table_data, month_name):
    """Записать данные в целевую таблицу SUMMARY"""
    try:
        # 1. Аутентификация
        gs = gspread.service_account('creds.json')

        # 2. Открыть целевую таблицу по ID
        target_spreadsheet = gs.open_by_key(
            '1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ')
        summary_sheet = target_spreadsheet.worksheet('SUMMARY')

        # 3. Получить текущие заголовки
        headers = summary_sheet.row_values(2)
        print(f"Заголовки в SUMMARY: {headers}")

        # 4. Нормализуем название месяца для сравнения
        normalized_month = month_name.capitalize()

        # 4. Найти столбец для месяца
        month_col = None
        for i, header in enumerate(headers, 1):  # Начинаем с 1 столбца
            if header == normalized_month:
                month_col = i
                print(
                    f"Найден столбец для месяца {normalized_month}: {header} (столбец {i})")
                break

        if month_col is None:
            # Найти первый пустой столбец
            for i, header in enumerate(headers, 1):
                if not header.strip():  # Пустой столбец
                    month_col = i
                    # Записать название месяца в вторую строку в ячейку номер month_col
                    summary_sheet.update_cell(2, month_col, normalized_month)
                    summary_sheet.update_cell(
                        3, month_col + 1, f"{normalized_month} %")
                    print(
                        f"Создан новый столбец для {normalized_month} в позиции: {month_col}")
                    break
        if month_col is None:
            # Добавить новые столбцы в конец
            month_col = len(headers) + 1
            if month_col > 37:  # Проверка ограничения Google Sheets
                print("✗ Достигнут лимит столбцов (37)")
                return False
            summary_sheet.update_cell(2, month_col, normalized_month)
            summary_sheet.update_cell(
                3, month_col + 1, f"{normalized_month} %")
            time.sleep(2)

        print(
            f"Столбец {normalized_month} найден/создан в позиции: {month_col}")

        # 5. Подготовить данные для записи
        update_data = []
        num_rows = len(table_data)
        # Начинаем с 2 строки в SUMMARY записываем данные из table_data которую получили из analyze
        # for i, (category, amount) in enumerate(table_data, start=4):
        #      update_data.append({
        #         'range': f"{gspread.utils.rowcol_to_a1(i, 1)}",
        #         'values': [[category]]
        #     })

        # for i, (category, amount, percentage) in enumerate(table_data, start=4):

        for i, row_data in enumerate(table_data, start=4):
            if len(row_data) == 3:
                category, amount, percentage = row_data
            elif len(row_data) == 2:
                category, amount = row_data
                percentage = 0  # или None, или пустая строка
            else:
                continue  # пропускаем некорректные строки

            # Запись категорий (только один раз)

            # update_data.append({
            #         'range': f'{gspread.utils.rowcol_to_a1(i, month_col +1)}',
            #         'values': [[percentage]]
            #     })

            # Запись суммы и процента
            update_data.append({
                'range': f"{gspread.utils.rowcol_to_a1(i, month_col)}",
                'values': [[amount]]
            })

            # update_data.append({
            #     'range': f"{gspread.utils.rowcol_to_a1(i, month_col+1)}",
            #     'values': [[percentage]]
            # })
            update_data.append({
                'range': f"{gspread.utils.rowcol_to_a1(i, month_col + 1)}",
                'values': [[percentage]]
            })

            # 6. Выполнить batch-запрос
        if update_data:
            batch_size = 10
            for i in range(0, len(update_data), batch_size):
                batch = update_data[i:i+batch_size]
            #     batch = [
            #         [
            #             t['category'],
            #             t['amount'],
            #             t['percentage']

            #         ] for t in table_data[i:i+batch_size]
            #     ]

                summary_sheet.batch_update(batch)
                if i + batch_size < len(update_data):
                    time.sleep(2)  # Пауза между пакетами

            print(
                f"✓ Данные за {normalized_month} записаны: {num_rows} строк")
            time.sleep(5)

        return True
    except Exception as e:
        print(f"✗ Ошибка записи в SUMMARY: {e}")
        return False


def main():
    transactions, daily_categories = load_transactions(FILE)
    if not transactions:
        print(f"No transactions found")
        return
    data = analyze(transactions, daily_categories)
    terminal_visualization(data)

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
            worksheet.delete_rows(1, len(all_values)+1)
        # worksheet.clear()
        time.sleep(2)

        # headers = ["Date", "Description", "Amount", "Type", "Category"]
        # worksheet.insert_row(headers, 7)

        # # Add transactions in batches
        # if transactions:
        #     batch_size = 9
        #     for i in range(0, len(transactions), batch_size):
        #         batch = [
        #             [
        #                 t['date'],
        #                 t['desc'][:50],
        #                 t['amount'],
        #                 t['type'],
        #                 t['category']
        #             ] for t in transactions[i:i+batch_size]
        #         ]
        #         worksheet.append_rows(batch)

        #         # No sleep after last batch
        #         if i + batch_size < len(transactions):
        #             time.sleep(10)
        # Подготовить все данные
        all_data = [["Date", "Description", "Amount", "Type", "Category"]]
        for t in transactions:
            all_data.append([t['date'], t['desc'][:50],
                            t['amount'], t['type'], t['category']])

        # Записать все данные одним запросом
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

        worksheet.format('B2:B4', {'numberFormat': {
            'type': 'CURRENCY', 'pattern': '€#,##0.00'}, "textFormat": {'bold': True, 'fontSize': 12}})
        time.sleep(2)
        worksheet.format('C8:C31', {'numberFormat': {
            'type': 'CURRENCY', 'pattern': '€#,##0.00'}})
        time.sleep(2)
        worksheet.format('A7:E7', {"textFormat": {
            'bold': True, 'fontSize': 12},
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}})
        time.sleep(2)

        worksheet.update('A2:A4', [['Total Income:'], [
            'Total Expenses:'], ['Savings:']])
        time.sleep(2)
        worksheet.update('B2:B4', [[total_income], [
            total_expense], [savings]])
        time.sleep(2)
        worksheet.update('C2:C4', [[1], [
            expense_rate], [savings_rate]])

        time.sleep(2)
        worksheet.format('C2:C4', {'numberFormat': {
            'type': 'PERCENT',
            'pattern': '0%'}})
        time.sleep(2)

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
            time.sleep(2)

        if category_data:

            last_row = 7 + len(category_data)
            last_row_transactions = 7 + len(transactions)
            table_data = []
            # for category, amount in sorted_categories:
            # percentage = (amount / total_expenses *
            #               100) if total_expenses > 0 else 0
            # table_data.append([category, amount, percentage/100])
            table_data = prepare_summary_data(data, transactions)
            time.sleep(2)

            # После подготовки table_data в main()
            print(f"Подготовлено строк данных: {len(table_data)}")
            # Проверяем, что у нас достаточно строк в таблице
            if last_row < 7 + len(table_data):
                # Добавляем нужное количество строк
                rows_to_add = (7 + len(table_data)) - last_row
                worksheet.add_rows(rows_to_add)
                print(f"Добавлено {rows_to_add} строк в таблицу")
                time.sleep(2)

            MONTH_NORMALIZED = get_month_column_name(
                MONTH)  # Нормализуем название месяца
            success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
            time.sleep(2)
            if success:
                print(
                    f"✓ Данные за {MONTH_NORMALIZED} успешно записаны в SUMMARY")
            else:
                print(f"✗ Ошибка при записи в SUMMARY")

            # range_str = 'G8:I{}'.format(last_row)
            # worksheet.update(range_str, category_data)

            category_headers = [['Category', 'Amount', 'Percentage']]
            worksheet.update('G7:I7', category_headers)
            time.sleep(2)
            # Подготавливаем данные для записи в правильном формате
            category_table_data = []
            for category, amount, percentage in table_data:
                if isinstance(percentage,(int, float)):
                    category_table_data.append([category, amount, percentage])
                else:
                    category_table_data.append([category, amount, 0])

            # Определяем правильный диапазон для записи
            end_row = 7 + len(category_table_data)
            # worksheet.update(f'G8:I{end_row}', category_table_data)
            # time.sleep(2)

            # Проверяем, достаточно ли строк в таблице
            current_rows = worksheet.row_count
            if end_row > current_rows:
                # Добавляем недостающие строки
                rows_to_add = end_row - current_rows
                worksheet.add_rows(rows_to_add)
                print(f"Добавлено {rows_to_add} строк в таблицу")
                time.sleep(2)

            # Записываем данные
            worksheet.update(f'G8:I{end_row}', category_table_data)
            time.sleep(2)

            # worksheet.update('G8:I7}', category_data)
            # worksheet.update(f'G8:I{last_row}', table_data)
            # time.sleep(2)

            # Определяем правильный диапазон для таблицы категорий
            category_end_row = 7 + len(table_data)

            # Проверяем, достаточно ли строк
            if category_end_row > worksheet.row_count:
                rows_to_add = category_end_row - worksheet.row_count
                worksheet.add_rows(rows_to_add)
                print(f"Добавлено {rows_to_add} строк для таблицы категорий")
                time.sleep(2)

            worksheet.update(f'G8:I{category_end_row}', table_data)
            time.sleep(2)

            worksheet.format('G7:I7', {
                "textFormat": {"bold": True, "fontSize": 12},
                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
            })
            worksheet.format(f'H8:H{end_row}',
                             {
                "numberFormat": {
                    "type": "CURRENCY",
                    "pattern": "€#,##0.00"
                }

            })
            time.sleep(2)
            worksheet.format(f'I8:I{end_row}', {
                "numberFormat": {
                    "type": "PERCENT",
                    "pattern": "0.00%"
                }
            })
            # worksheet.format(f'I9', {
            #     "numberFormat": {
            #         "type": "PERCENT",
            #         "pattern": "0.00%"
            #     }
            # })
            time.sleep(2)
            column_formats = [
                (f'A8:A{last_row_transactions}', {"backgroundColor": {
                 "red": 0.90, "green": 0.90, "blue": 0.90}}),  # Очень светло-серый
                (f'B8:B{last_row_transactions}', {"backgroundColor": {
                 "red": 0.96, "green": 0.96, "blue": 0.96}}),
                (f'C8:C{last_row_transactions}', {"backgroundColor": {
                 "red": 0.94, "green": 0.94, "blue": 0.94}}),
                (f'D8:D{last_row_transactions}', {"backgroundColor": {
                 "red": 0.92, "green": 0.92, "blue": 0.92}}),
                (f'E8:E{last_row_transactions}', {"backgroundColor": {
                 "red": 0.90, "green": 0.90, "blue": 0.90}})  # Более темный серый
            ]
            # Применяем форматирование заголовков
            # worksheet.format('A8:E8', header_format)
            for range_, format_ in column_formats:
                worksheet.format(range_, format_)
            time.sleep(2)
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
            time.sleep(2)

            worksheet.update('A2:A4', [['Total Income:'], [
                             'Total Expenses:'], ['Savings:']])
            time.sleep(2)

            worksheet.update('B2:B4', [[total_income], [
                             total_expense], [savings]])
            time.sleep(2)
            worksheet.update('C2:C4', [[1], [
                             expense_rate], [savings_rate]])
            time.sleep(2)
            worksheet.format('C2:C4', {'numberFormat': {
                'type': 'PERCENT',
                'pattern': '0%'}})

            border_style = {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0.6, "green": 0.6, "blue": 0.6}
            }

            border_format = {
                "borders": {
                    "top": border_style,
                    "bottom": border_style,
                    "left": border_style,
                    "right": border_style
                }
            }

            tables = [
                f'A7:E{7 + len(transactions)}',  # Основная таблица транзакций
                f'G7:I{end_row}',  # Таблица категорий
                'A2:C4'                          # Блок с итогами
            ]
            for table_range in tables:
                worksheet.format(table_range, border_format)
            time.sleep(2)

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
            time.sleep(2)
            worksheet.format('G7:I7', header_bottom_border)
            time.sleep(2)
            worksheet.format('D2:D4', header_left_border)
            time.sleep(2)
            worksheet.format('A2:C4', border_format)
            time.sleep(2)
            # worksheet.format('A2:C2', header_bottom_border)
            worksheet.format('A1:C1', header_bottom_border)
            time.sleep(2)
            worksheet.format('A4:C4', border_format)
            time.sleep(2)
            worksheet.format('A5:C5', header_top_border)
            time.sleep(2)

            recommendations = generate_daily_recommendations(data)
            time.sleep(2)

            rec_headers = ["Priority", "Recommendation"]
            rec_data = [[f"{i+1}.", rec]
                        for i, rec in enumerate(recommendations)]

            rec_start_row = 7
            rec_start_col = 11

            worksheet.update(
                values=[rec_headers],
                # range_name = f"{gspread.utils.rowcol_to_a1(rec_start_row, rec_start_col)}:{gspread.utils.rowcol_to_a1(rec_start_row, rec_start_col + 1)}"
                range_name=f"K{rec_start_row}:L{rec_start_row}"


            )
            time.sleep(2)
            for i, row in enumerate(rec_data, start=rec_start_row+1):
                worksheet.update(f"K{i}:L{i}", [row])
            time.sleep(2)
            fmt = cellFormat(
                horizontalAlignment='CENTER',
                padding=Padding(top=8, right=12, bottom=8, left=12),
                wrapStrategy='WRAP'
            )
            time.sleep(2)
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
            time.sleep(2)

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
            time.sleep(2)
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
            time.sleep(2)

            worksheet.update(f"K6", [['DAILY RECOMMENDATIONS']])
            time.sleep(2)
            worksheet.format("K6", {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            time.sleep(2)
            worksheet.merge_cells(f"K6:L6")
            time.sleep(2)

            worksheet.update(f"A6", [['FINANCIAL OVERVIEW']])
            time.sleep(2)
            worksheet.format("A6", {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            time.sleep(2)
            worksheet.merge_cells(f"A6:E6")
            time.sleep(2)
            worksheet.update(f"G6", [['TRANSACTION CATEGORIES']])
            time.sleep(2)
            worksheet.format("G6", {
                "textFormat": {"bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER"
            })
            time.sleep(2)
            worksheet.merge_cells(f"G6:I6")
            time.sleep(2)
            print(
                f"Added {len(recommendations)} recommendations to Google Sheets")
            worksheet.format("A1:Z100", {"horizontalAlignment": "CENTER"})
            time.sleep(2)

            # Записать данные
            MONTH_NORMALIZED = get_month_column_name(MONTH)
            success = write_to_target_sheet(table_data, MONTH_NORMALIZED)
            if success:
                print(f"✓ Данные за {MONTH} успешно записаны в SUMMARY")
            else:
                print(f"✗ Ошибка при записи в SUMMARY")


# Authenticate and open Google Sheets
        # gs = gspread.service_account('creds.json')
        # target_spreadsheet = gspread.service_account('creds.json').open_by_key(
        #     '1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ')

        # # Check if worksheet exists
        # worksheet = None
        # try:
        #     worksheet = target_spreadsheet.worksheet(MONTH)
        #     print(f"\nWorksheet '{MONTH}' found. Updating...")
        # except gspread.WorksheetNotFound:
        #     print(f"\nWorksheet for {MONTH} not found. Creating a new one...")
        #     # First check if we've reached the sheet limit (max 200 sheets)
        #     if len(sh.worksheets()) >= 200:
        #         raise Exception("Maximum number of sheets (200) reached")

        #     # Check if sheet exists but with different case (e.g. "march" vs "March")
        #     existing_sheets = [ws.title for ws in sh.worksheets()]
        #     if MONTH.lower() in [sheet.lower() for sheet in existing_sheets]:
        #         # Find the existing sheet with case-insensitive match
        #         for sheet in sh.worksheets():
        #             if sheet.title.lower() == MONTH.lower():
        #                 worksheet = sheet
        #                 print(
        #                     f"Using existing worksheet '{sheet.title}' (case difference)")
        #                 break
        #     else:
        #         # Create new worksheet with unique name if needed
        #         try:
        #             worksheet = target_spreadsheet.add_worksheet(
        #                 title=MONTH, rows="100", cols="20")
        #             print(f"New worksheet '{MONTH}' created successfully.")
        #         except gspread.exceptions.APIError as e:
        #             if "already exists" in str(e):
        #                 # If we get here, it means the sheet exists but wasn't found earlier
        #                 worksheet = target_spreadsheet.worksheet(MONTH)
        #                 print(f"Worksheet '{MONTH}' exists. Using it.")
        #             else:
        #                 raise e

        # if worksheet is None:
        #     raise Exception("Failed to access or create worksheet")

        # Clear existing data (keep headers)
        # all_values = worksheet.get_all_values()
        # if len(all_values) > 1:
        #     worksheet.delete_rows(1, len(all_values)+1)

        #     headers = ["Date", "Description", "Amount", "Type", "Category"]
        #     worksheet.insert_row(headers, 7)

        # Add transactions in batches

        # if transactions:
        #     batch_size = 12
        #     for i in range(0, len(transactions), batch_size):
        #         batch = [
        #             [
        #                 t['date'],
        #                 t['desc'][:50],
        #                 t['amount'],
        #                 t['type'],
        #                 t['category']
        #             ] for t in transactions[i:i+batch_size]
        #         ]
        #         worksheet.append_rows(batch)

        #         # No sleep after last batch
        #         if i + batch_size < len(transactions):
        #             time.sleep(3)

        #     total_income = sum(t['amount']
        #                        for t in transactions if t['type'] == 'income')
        #     total_expense = sum(t['amount']
        #                         for t in transactions if t['type'] == 'expense')
        #     savings = total_income - total_expense

            print(
                f"\nSuccessfully updated {len(transactions)} transactions in Google Sheets")
        else:
            print("\nNo transactions to update in Google Sheets")

    except Exception as e:
        print(f"\nError in Google Sheets operation: {str(e)}")


if __name__ == "__main__":
    main()
