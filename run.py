import csv
import gspread
import time
import warnings
import os
import sys
import json
import threading
import tempfile
import shutil
from collections import defaultdict
from datetime import datetime
from gspread_formatting import *
from gspread.utils import rowcol_to_a1
from google.oauth2 import service_account
from flask import Flask, request, render_template_string
from werkzeug.utils import secure_filename

warnings.filterwarnings("ignore", category=DeprecationWarning)
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-123')

OPERATION_STATUS = {}

DAILY_NORMS = {
    "Rent": 50.0,
    "Gym": 3.0,
    "Groceries": 3,
    "Transport": 0.27,
    "Entertainment": 0.17,
    "Utilities": 2.0,
    "Shopping": 3.33,
    "Dining": 10.00,
}

ALLOWED_EXTENSIONS = {"csv", "txt"}


def allowed_file(filename):
    """
    Проверяет, что файл имеет допустимое расширение и имя.
    Возвращает True если файл допустим, False в противном случае.
    """
    # Проверяем основные условия
    if not filename or not isinstance(filename, str):
        return False
    
    # Запрещаем скрытые файлы (начинающиеся с точки)
    if filename.startswith('.'):
        return False
    
    # Проверяем наличие точки в имени файла
    if '.' not in filename:
        return False
    
    # Проверяем длину имени файла (не более 100 символов)
    if len(filename) > 100:
        return False
    
    # Извлекаем расширение и проверяем его
    try:
        extension = filename.rsplit('.', 1)[1].lower()
    except IndexError:
        return False
    
    # Проверяем, что расширение в списке допустимых
    return extension in ALLOWED_EXTENSIONS  # Ограничение длины имени


def get_google_credentials():
    """Get Google credentials with better error handling"""
    try:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        if "DYNO" in os.environ:
            print("🔑 Using environment credentials from Heroku")
            service_account_json = os.environ.get(
                "GOOGLE_SERVICE_ACCOUNT_JSON"
            )
            if service_account_json:
                try:
                    creds_dict = json.loads(service_account_json)
                    Credentials = service_account.Credentials
                    creator = Credentials.from_service_account_info
                    return creator(creds_dict, scopes=SCOPES)
                except json.JSONDecodeError:
                    print("❌ Invalid JSON in GOOGLE_SERVICE_ACCOUNT_JSON")
                    return None
            else:
                print(
                    "❌ GOOGLE_SERVICE_ACCOUNT_JSON environment  "
                    "variable not found"
                )
                return None
        else:
            # Local development
            if os.path.exists("creds.json"):
                Credentials = service_account.Credentials
                credentials = Credentials.from_service_account_file(
                    "creds.json", scopes=SCOPES
                )
                return credentials
            else:
                print("❌ Local creds.json file not found")
                print(
                    "💡 Create creds.json  "
                    "with Google Service Account credentials"
                )
                return None
    except Exception as e:
        print(f"❌ Error getting credentials: {e}")
        return None


def categorize(description):
    """Categorize transaction based on description."""
    desc = description.lower()

    categories = {
        "Salary": ["salary", "wages", "salary deposit"],
        "Bonus": ["bonus", "tip", "reward"],
        "Rent": ["rent", "monthly rent", "rent payment"],
        "Groceries": ["supermarket", "grocery", "food"],
        "Dining": ["restaurant", "cafe", "coffee"],
        "Transport": ["bus", "train", "taxi", "uber", "fuel", "enoc"],
        "Entertainment": ["movie", "netflix", "concert", "spotify"],
        "Utilities": ["electricity", "water", "gas", "internet", "phone"],
        "Gym": ["gym", "fitness", "yoga"],
        "Shopping": ["clothing", "electronics", "shopping"],
        "Health": ["pharmacy", "doctor", "health", "dentist"],
        "Insurance": ["insurance", "health insurance", "car insurance"],
        "Travel": ["flight", "hotel", "travel", "airline", "hilton"],
        "Other": [],
    }

    for cat, terms in categories.items():
        if any(term in desc for term in terms):
            return cat
    return "Other"


def get_month_column_name(month_input):
    """Convert month name to standard format"""
    month = month_input.strip().capitalize()
    month_mapping = {
        "Jan": "January",
        "Feb": "February",
        "Mar": "March",
        "Apr": "April",
        "May": "May",
        "Jun": "June",
        "Jul": "July",
        "Aug": "August",
        "Sep": "September",
        "Oct": "October",
        "Nov": "November",
        "Dec": "December",
    }
    return month_mapping.get(month, month)


def analyze(transactions, daily_categories, month):
    """Perform financial analysis with daily tracking"""
    analysis = {
        "income": 0,
        "expenses": 0,
        "categories": defaultdict(float),
        "income_categories": defaultdict(float),
        "month": month,
        "daily_categories": daily_categories,
        "days_count": 30,
        "daily_averages": defaultdict(float),
        "norms_violations": [],
    }

    for t in transactions:
        if t["type"] == "income":
            analysis["income"] += t["amount"]
            analysis["income_categories"][t["category"]] += t["amount"]
        else:
            analysis["expenses"] += abs(t["amount"])
            analysis["categories"][t["category"]] += abs(t["amount"])

    # Calculate daily averages
    for category, total in analysis["categories"].items():
        daily_avg = total / analysis["days_count"]
        analysis["daily_averages"][category] = daily_avg
        if category in DAILY_NORMS:
            if daily_avg > DAILY_NORMS[category] * 1.1:
                analysis["norms_violations"].append(
                    f"Daily average for {category} "
                    f"overspent: {daily_avg:.2f}€ "
                    f"vs norm: {DAILY_NORMS[category]:.2f}€"
                )

    analysis["savings"] = analysis["income"] - analysis["expenses"]
    return analysis


def format_terminal_output(data, month, transactions_count=0):
    """Format terminal output for 80x24 characters as in screenshot"""
    output = []
    output.append(" ")

    expense_rate = (
        (data["expenses"] / data["income"] * 100) if data["income"] > 0 else 0
    )

    savings_rate = (
        (data["savings"] / data["income"] * 100) if data["income"] > 0 else 0
    )

    centered_title = f"<u>FINANCIAL OVERVIEW: {month.upper()}</u>"
    output.append(centered_title)
    output.append(f"Income: {data['income']:8.2f}€ [{'■' * 20}] 100.0%")
    output.append(
        f"Expenses: {data['expenses']:8.2f}€ "
        f"[{'■' * int(expense_rate / 5)}] "
        f"{expense_rate:.1f}%"
    )
    output.append(
        f"Savings: {data['savings']:8.2f}€ "
        f"[{'■' * int(savings_rate / 5)}] "
        f"{savings_rate:.1f}%"
    )

    output.append("<u>EXPENSE CATEGORIES:</u>")
    top_categories = sorted(
        data["categories"].items(), key=lambda x: x[1], reverse=True
    )[:12]

    for category, amount in top_categories:
        if data["expenses"] > 0:
            percent = amount / data["expenses"] * 100
            bar_length = max(1, int(percent / 5))
            output.append(
                f"{category[:15]:<15} {amount:8.2f}€ "
                f"{'■' * bar_length} ({percent:.1f}%)"
            )
        else:
            output.append(f"{category[:15]:<15} {amount:8.2f}€")

    output.append("<u>DAILY SPENDING and NORMS:</u>")
    sorted_categories = sorted(
        [
            (cat, avg)
            for cat, avg in data["daily_averages"].items()
            if cat in DAILY_NORMS
        ],
        key=lambda x: abs(x[1] - DAILY_NORMS.get(x[0], 0)),
        reverse=True,
    )[:3]

    for category, avg in sorted_categories:
        norm = DAILY_NORMS.get(category, 0)
        diff = avg - norm
        arrow = "▲" if diff > 0 else "▼"
        output.append(
            f"{category[:12]:<12} Avg: {avg:5.2f}€ "
            f"Norm: {norm:5.2f}€ {arrow} {abs(diff):.2f}€"
        )

    output.append("<u>DAILY SPENDING RECOMMENDATIONS:</u>")
    recommendations = generate_daily_recommendations(data)[:3]
    for i, rec in enumerate(recommendations, 1):
        if len(rec) > 70:
            rec = rec[:67] + "..."
        output.append(f"{i}. {rec}")

    while len(output) > 24:
        output.pop()

    return "\n".join(output)


def terminal_visualization(data):
    """Visualize financial data in terminal (80x24) as in screenshot."""
    centered_title = f"FINANCIAL OVERVIEW {data['month'].upper()}"
    print(centered_title)

    print(f"Income:   {data['income']:8.2f}€ [{'■' * 20}] 100.0%")

    expense_rate = (
        (data["expenses"] / data["income"] * 100) if data["income"] > 0 else 0
    )

    savings_rate = (
        (data["savings"] / data["income"] * 100) if data["income"] > 0 else 0
    )

    print(
        f"Expenses: {data['expenses']:8.2f}€ "
        f"[{'■' * int(expense_rate / 5)}] "
        f"{expense_rate:.1f}%"
    )

    print(
        f"Savings:  {data['savings']:8.2f}€"
        f"[{'■' * int(savings_rate / 5)}] "
        f"{savings_rate:.1f}%"
    )

    print("EXPENSE CATEGORIES: ")
    top_categories = sorted(
        data["categories"].items(), key=lambda x: x[1], reverse=True
    )[:8]

    categories_with_percent = []
    max_percent = (
        max(
            (amount / data["expenses"] * 100)
            for category, amount in top_categories
        )
        if data["expenses"] > 0
        else 0
    )

    for category, amount in top_categories:
        percent = (
            (amount / data["expenses"] * 100) if data["expenses"] > 0 else 0
        )

        if max_percent > 0:
            scaled_percent = max(1, int(percent / max_percent * 8))
        else:
            scaled_percent = 1

        categories_with_percent.append((category, amount, scaled_percent))

    col1 = categories_with_percent[0:3]
    col2 = categories_with_percent[3:6]
    col3 = categories_with_percent[6:9]

    for i in range(3):
        line = ""
        if i < len(col1):
            cat1, amt1, bar_len1 = col1[i]
            line += f"{cat1[:10]:<10} {amt1:6.2f}€ {'■' * bar_len1}"
        else:
            line += " " * 25

        line += " " * 2

        if i < len(col2):
            cat2, amt2, bar_len2 = col2[i]
            line += f"{cat2[:10]:<10} {amt2:6.2f}€ {'■' * bar_len2}"
        else:
            line += " " * 25

        line += " " * 2

        if i < len(col3):
            cat3, amt3, bar_len3 = col3[i]
            line += f"{cat3[:10]:<10} {amt3:6.2f}€ {'■' * bar_len3}"

        print(line)

    print("DAILY SPENDING and NORMS: ")
    sorted_categories = sorted(
        [
            (cat, avg)
            for cat, avg in data["daily_averages"].items()
            if cat in DAILY_NORMS
        ],
        key=lambda x: x[1] - DAILY_NORMS.get(x[0], 0),
        reverse=True,
    )[:3]

    for category, avg in sorted_categories:
        norm = DAILY_NORMS.get(category, 0)
        diff = avg - norm
        arrow = "▲" if diff > 0 else "▼"
        print(
            f"{category[:10]:<10}"
            f"Avg: {avg:5.2f}€ Norm: {norm:5.2f}€"
            f"{arrow} {abs(diff):.2f}€"
        )

    print("DAILY SPENDING RECOMMENDATIONS: ")
    recommendations = generate_daily_recommendations(data)[:3]
    for i, rec in enumerate(recommendations, 1):
        if len(rec) > 70:
            rec = rec[:67] + "..."
        print(f"{i}. {rec}")


def generate_daily_recommendations(data):
    """Generate daily category-specific recommendations."""
    recs = []

    if not data or "income" not in data:
        return ["No financial data available for recommendations."]

    if data["income"] <= 0:
        return ["No income data - cannot generate recommendations."]
    else:
        # 1. Savings rate recommendation
        expense_rate = data["expenses"] / data["income"] * 100
        savings_rate = data["savings"] / data["income"] * 100

        if savings_rate < 20:
            recs.append(f"Aim for 20% savings (current: {savings_rate:.1f}%)")
            recs.extend(data["norms_violations"][:3])

        # Ensure minimum recommendations
        if len(recs) < 3:
            recs.extend(
                [
                    "Plan meals weekly to reduce grocery costs",
                    "Use public transport more frequently",
                ]
            )

        return recs[:3]


def prepare_summary_data(data, transactions):
    """Prepare the data for the SUMMARY section - all categories and totals."""
    all_categories = [
        "TOTAL INCOME",
        "TOTAL EXPENSES",
        "SAVINGS",
        "",
        "INCOME CATEGORIES:",
        "Salary",
        "Bonus",
        "Other Income",
        "",
        "EXPENSE CATEGORIES:",
        "Rent",
        "Groceries",
        "Dining",
        "Transport",
        "Entertainment",
        "Utilities",
        "Gym",
        "Shopping",
        "Health",
        "Insurance",
        "Education",
        "Travel",
        "Car",
        "Other",
    ]

    # Collecting data by income
    income_by_category = defaultdict(float)
    for t in transactions:
        if t["type"] == "income":
            income_by_category[t["category"]] += t["amount"]

    # Collecting data by expense
    expenses_by_category = defaultdict(float)
    for t in transactions:
        if t["type"] == "expense":
            expenses_by_category[t["category"]] += t["amount"]

    # Preparing totals
    table_data = []
    for category in all_categories:
        if category == "TOTAL INCOME":
            table_data.append([category, data["income"], 1.0])
        elif category == "TOTAL EXPENSES":
            percentage = (
                data["expenses"] / data["income"] if data["income"] > 0 else 0
            )
            table_data.append([category, data["expenses"], percentage])
        elif category == "SAVINGS":
            percentage = (
                data["savings"] / data["income"] if data["income"] > 0 else 0
            )
            table_data.append([category, data["savings"], percentage])
        elif category in ["", "INCOME CATEGORIES:", "EXPENSE CATEGORIES:"]:
            table_data.append([category, "", ""])
        elif category in income_by_category:
            amount = income_by_category[category]
            percentage = amount / data["income"] if data["income"] > 0 else 0
            table_data.append([category, amount, percentage])
        elif category == "Salary":
            matched = False
            for income_cat in income_by_category:
                amount = income_by_category[income_cat]
                percentage = (
                    amount / data["income"] if data["income"] > 0 else 0
                )
                table_data.append([category, amount, percentage])
                matched = True
                break

            if not matched:
                if category == "Salary":
                    for income_cat in income_by_category:
                        if (
                            "salary" in income_cat.lower()
                            or "income" in income_cat.lower()
                        ):
                            amount = income_by_category[income_cat]
                            percentage = (
                                amount / data["income"]
                                if data["income"] > 0
                                else 0
                            )
                            table_data.append([category, amount, percentage])
                            matched = True
                            break

            if not matched:
                table_data.append([category, 0, 0])
        elif category in expenses_by_category:
            amount = expenses_by_category[category]
            percentage = (
                amount / data["expenses"] if data["expenses"] > 0 else 0
            )
            table_data.append([category, amount, percentage])
        else:
            table_data.append([category, 0, 0])

    return table_data


def set_column_width(worksheet, column_letter, width):
    """Set column width for worksheet"""
    try:
        # Convert column letter to index (A=1, B=2, etc.)
        col_index = gspread.utils.a1_to_rowcol(column_letter + "1")[1]

        body = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": col_index - 1,
                            "endIndex": col_index,
                        },
                        "properties": {"pixelSize": width},
                        "fields": "pixelSize",
                    }
                }
            ]
        }
        worksheet.spreadsheet.batch_update(body)
    except Exception as e:
        print(f"⚠️ Error setting column width for {column_letter}: {e}")


def write_to_month_sheet(month_name, transactions, data):
    """Write data to month worksheet in Google Sheets"""
    try:
        print(f"📊 Writing to {month_name} worksheet...")

        # 1. Authentication
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials for month sheet")
            return False

        gc = gspread.authorize(creds)
        sh = gc.open("Personal Finances")

        # 2. Get or create worksheet
        try:
            worksheet = sh.worksheet(month_name)
        except gspread.WorksheetNotFound:
            print(f"📝 Creating new worksheet '{month_name}'...")
            worksheet = sh.add_worksheet(
                title=month_name, rows="100", cols="20"
            )
            print(f"✅ Worksheet '{month_name}' created")
            time.sleep(3)

        # 3. Clear existing data
        try:
            worksheet.clear()
            time.sleep(3)
        except Exception as e:
            print(f"⚠️ Warning: Could not clear worksheet: {e}")

        # 4. Main data with retries
        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            try:
                # Financial Overview header
                worksheet.update("A6", [["FINANCIAL OVERVIEW"]])
                worksheet.merge_cells("A6:E6")

                # Table headers
                headers = ["Date", "Description", "Amount", "Type", "Category"]
                worksheet.update("A7", [headers])

                # Write transactions
                all_data = []
                for t in transactions:
                    all_data.append(
                        [
                            t["date"],
                            t["desc"][:30],
                            t["amount"],
                            t["type"],
                            t["category"],
                        ]
                    )

                if all_data:
                    worksheet.update("A8", all_data)

                # Transaction Categories header
                worksheet.update("G6", [["TRANSACTION CATEGORIES"]])
                worksheet.merge_cells("G6:I6")

                # Categories table headers
                category_headers = ["Category", "Amount", "Percentage"]
                worksheet.update("G7", [category_headers])

                # Prepare and write category data
                table_data = prepare_summary_data(data, transactions)
                category_data = []
                for row in table_data:
                    if row[0] and row[0] not in [
                        "",
                        "INCOME CATEGORIES:",
                        "EXPENSE CATEGORIES:",
                    ]:
                        category_data.append([row[0], row[1], row[2]])

                if category_data:
                    worksheet.update("G8", category_data)

                # Daily Recommendations header
                worksheet.update("K6", [["DAILY RECOMMENDATIONS"]])
                worksheet.merge_cells("K6:L6")

                # Recommendations headers
                rec_headers = ["Priority", "Recommendation"]
                worksheet.update("K7", [rec_headers])

                # Write recommendations
                recommendations = generate_daily_recommendations(data)
                rec_data = []
                for i, rec in enumerate(recommendations, 1):
                    rec_data.append([f"{i}", rec[:100]])

                if rec_data:
                    worksheet.update("K8", rec_data)

                # Summary section at the top
                expense_percentage = (
                    (data["expenses"] / data["income"])
                    if data["income"] > 0
                    else 0
                )
                savings_percentage = (
                    (data["savings"] / data["income"])
                    if data["income"] > 0
                    else 0
                )

                summary_data = [
                    ["Total Income:", data["income"], 1.0],
                    ["Total Expenses:", data["expenses"], expense_percentage],
                    ["Savings:", data["savings"], savings_percentage],
                ]
                worksheet.update("A2", summary_data)

                break

            except Exception as e:
                retry_count += 1
                if "429" in str(e) or "Quota exceeded" in str(e):
                    wait_time = 60 * retry_count
                    print(
                        f"⚠️ Rate limit exceeded. "
                        f"Retry {retry_count}/{max_retries} "
                        f"in {wait_time} seconds..."
                    )
                    time.sleep(wait_time)
                else:
                    print(f"❌ Error writing data: {e}")
                    raise e

        if retry_count >= max_retries:
            print(f"❌ Failed to write data after {max_retries} retries")
            return False

        # 5. Formatting with colors
        try:
            last_transaction_row = 7 + len(transactions)
            last_category_row = 7 + len(category_data) if category_data else 7
            last_rec_row = 7 + len(rec_data) if rec_data else 7

            # Financial Overview header
            worksheet.format(
                "A6",
                {
                    "textFormat": {"bold": True, "fontSize": 14},
                    "horizontalAlignment": "CENTER",
                },
            )

            # Table headers
            worksheet.format(
                "A7:E7",
                {
                    "textFormat": {"bold": True, "fontSize": 12},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1},
                    },
                },
            )

            # Transaction table with alternating colors
            if last_transaction_row > 7:
                worksheet.format(
                    f"A8:A{last_transaction_row}",
                    {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.95,
                            "blue": 0.95,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                    },
                )

                worksheet.format(
                    f"B8:B{last_transaction_row}",
                    {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                    },
                )

                worksheet.format(
                    f"C8:C{last_transaction_row}",
                    {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.95,
                            "blue": 0.95,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                        "numberFormat": {
                            "type": "CURRENCY",
                            "pattern": "€#,##0.00",
                        },
                    },
                )

                worksheet.format(
                    f"D8:D{last_transaction_row}",
                    {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                    },
                )

                worksheet.format(
                    f"E8:E{last_transaction_row}",
                    {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.95,
                            "blue": 0.95,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                    },
                )

            # Transaction Categories header
            worksheet.format(
                "G6",
                {
                    "textFormat": {"bold": True, "fontSize": 14},
                    "horizontalAlignment": "CENTER",
                },
            )

            # Categories table headers
            worksheet.format(
                "G7:I7",
                {
                    "textFormat": {"bold": True, "fontSize": 12},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1},
                    },
                },
            )

            # Category table with alternating colors
            if last_category_row > 7:
                worksheet.format(
                    f"G8:G{last_category_row}",
                    {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.95,
                            "blue": 0.95,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                    },
                )

                worksheet.format(
                    f"H8:H{last_category_row}",
                    {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                        "numberFormat": {
                            "type": "CURRENCY",
                            "pattern": "€#,##0.00",
                        },
                    },
                )

                worksheet.format(
                    f"I8:I{last_category_row}",
                    {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.95,
                            "blue": 0.95,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                        "numberFormat": {
                            "type": "PERCENT",
                            "pattern": "0.00%",
                        },
                    },
                )

            # Daily Recommendations header
            worksheet.format(
                "K6",
                {
                    "textFormat": {"bold": True, "fontSize": 14},
                    "horizontalAlignment": "CENTER",
                },
            )

            # Recommendations headers
            worksheet.format(
                "K7:L7",
                {
                    "textFormat": {"bold": True, "fontSize": 12},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1},
                    },
                },
            )

            # Recommendations table with alternating colors
            if last_rec_row > 7:
                worksheet.format(
                    f"K8:K{last_rec_row}",
                    {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.95,
                            "blue": 0.95,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                    },
                )

                worksheet.format(
                    f"L8:L{last_rec_row}",
                    {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0,
                        },
                        "borders": {
                            "top": {"style": "SOLID", "width": 1},
                            "bottom": {"style": "SOLID", "width": 1},
                            "left": {"style": "SOLID", "width": 1},
                            "right": {"style": "SOLID", "width": 1},
                        },
                        "wrapStrategy": "WRAP",
                    },
                )

            # Summary section formatting
            worksheet.format(
                "A2:A4",
                {
                    "textFormat": {"bold": True},
                    "backgroundColor": {
                        "red": 0.95,
                        "green": 0.95,
                        "blue": 0.95,
                    },
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1},
                    },
                },
            )

            worksheet.format(
                "B2:B4",
                {
                    "textFormat": {"bold": False},
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1},
                    },
                    "numberFormat": {
                        "type": "CURRENCY",
                        "pattern": "€#,##0.00",
                    },
                },
            )

            worksheet.format(
                "C2:C4",
                {
                    "textFormat": {"bold": False},
                    "backgroundColor": {
                        "red": 0.95,
                        "green": 0.95,
                        "blue": 0.95,
                    },
                    "borders": {
                        "top": {"style": "SOLID", "width": 1},
                        "bottom": {"style": "SOLID", "width": 1},
                        "left": {"style": "SOLID", "width": 1},
                        "right": {"style": "SOLID", "width": 1},
                    },
                    "numberFormat": {"type": "PERCENT", "pattern": "0.00%"},
                },
            )

        except Exception as format_error:
            print(f"⚠️ Formatting error: {format_error}")

        # 6. Set column widths
        try:
            set_column_width(worksheet, "A", 100)
            set_column_width(worksheet, "B", 200)
            set_column_width(worksheet, "C", 80)
            set_column_width(worksheet, "D", 80)
            set_column_width(worksheet, "E", 100)
            set_column_width(worksheet, "G", 150)
            set_column_width(worksheet, "H", 80)
            set_column_width(worksheet, "I", 100)
            set_column_width(worksheet, "K", 90)
            set_column_width(worksheet, "L", 300)

        except Exception as width_error:
            print(f"⚠️ Column width error: {width_error}")

        return True

    except Exception as e:
        print(f"❌ Error writing to {month_name} worksheet: {e}")
        import traceback

        print(f"🔍 Traceback: {traceback.format_exc()}")
        return False


def sync_google_sheets_operation(month_name, table_data):
    """Synchronous Google Sheets operation"""
    try:

        # 1. Authentication
        creds = get_google_credentials()
        if not creds:
            print("❌ No credentials available")
            return False

        gc = gspread.authorize(creds)

        # 2. Open target spreadsheet by ID
        try:
            spreadsheet_key = "1US65_F99qrkqbl2oVkMa4DGUiLacEDRoNz_J9hr2bbQ"
            target_spreadsheet = gc.open_by_key(spreadsheet_key)
        except Exception as e:
            print(f"❌ Error opening spreadsheet: {e}")
            return False

        try:
            summary_sheet = target_spreadsheet.worksheet("SUMMARY")
        except Exception as e:
            print(f"❌ Error accessing SUMMARY worksheet: {e}")
            return False

        # 3. Get current headers
        headers = summary_sheet.row_values(2)

        # 4. Normalizing month name for comparison
        normalized_month = month_name.capitalize()

        # 5. Find the month column
        month_col = None
        for i, header in enumerate(headers, 1):
            if header == normalized_month:
                month_col = i
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
                    summary_sheet.update_cell(
                        3, month_col + 1, f"{normalized_month} %"
                    )
                    print(
                        f"✅ Created new column for {normalized_month}  "
                        f"at position: {month_col}"
                    )
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
            summary_sheet.update_cell(
                3, month_col + 1, f"{normalized_month} %"
            )
            print(
                f"✅ Added new column for {normalized_month} "
                f"at position: {month_col}"
            )

        # 6. Prepare data to be written
        update_data = []
        for i, row_data in enumerate(table_data, start=4):
            if len(row_data) == 3:
                category, amount, percentage = row_data
                update_data.append(
                    {
                        "range": f"{rowcol_to_a1(i, month_col)}",
                        "values": [[amount]],
                    }
                )
                update_data.append(
                    {
                        "range": f"{rowcol_to_a1(i, month_col + 1)}",
                        "values": [[percentage]],
                    }
                )

        # 7. Batch update
        if update_data:
            batch_size = 2
            max_retries = 3

            for i in range(0, len(update_data), batch_size):
                batch = update_data[i: i + batch_size]
                retry_count = 0
                success = False

                while not success and retry_count < max_retries:
                    try:
                        summary_sheet.batch_update(batch)
                        success = True

                    except Exception as e:
                        if "429" in str(e) or "Quota exceeded" in str(e):
                            retry_count += 1
                            wait_time = 60 * retry_count
                            print(
                                f"⚠️ Rate limit exceeded. "
                                f"Retry {retry_count}/{max_retries} "
                                f"in {wait_time} seconds..."
                            )
                            time.sleep(wait_time)
                        else:
                            print(f"❌ Error in batch update: {e}")
                            raise e

                if not success:
                    print(
                        f"❌ Failed to write batch {
                            i // batch_size + 1} after {max_retries} retries"
                    )
                    return False

                if i + batch_size < len(update_data):
                    time.sleep(15)

        return True

    except Exception as e:
        print(f"❌ Error in sync_google_sheets_operation: {e}")
        import traceback

        print(f"🔍 Traceback: {traceback.format_exc()}")
        return False


def write_to_target_sheet(table_data, month_name):
    """Write data to target SUMMARY sheet"""
    try:
        if not table_data:
            print("✗ No data to write to target sheet")
            return False

        if len(table_data) > 50:
            print(
                f"⚠️ Large dataset({
                    len(table_data)} rows), simplifying update"
            )
            simplified_data = []
            for row in table_data:
                if row[0] in ["TOTAL INCOME", "TOTAL EXPENSES", "SAVINGS"]:
                    simplified_data.append(row)
                elif row[0] and not any(
                    x in row[0] for x in ["CATEGORIES", ""]
                ):
                    simplified_data.append([row[0], row[1], 0])
            table_data = simplified_data

        return sync_google_sheets_operation(month_name, table_data)

    except Exception as e:
        print(f"✗ Error in writing into SUMMARY: {e}")
        return False


# def load_transactions(file_path_or_object):
#     """Load transactions from uploaded file with proper CSV parsing"""
#     transactions = []
#     daily_categories = defaultdict(lambda: defaultdict(float))

#     try:
#         # Handle both file objects and file paths
#         if hasattr(file_path_or_object, "read"):
#             # File object - read content
#             # Rewind file to beginning for mobile devices
#             file_path_or_object.seek(0)
#             content = file_path_or_object.read()
#             if isinstance(content, bytes):
#                 content = content.decode("utf-8")
#             lines = content.split("\n")
#         else:
#             # File path
#             with open(file_path_or_object, "r", encoding="utf-8") as file:
#                 lines = file.readlines()

#         # Parse CSV lines
#         for line in lines:
#             line = line.strip()
#             if line and not line.startswith(("#", "Date")):
#                 try:
#                     parts = line.split(",")
#                     if len(parts) >= 5:
#                         # Parse date (assuming format: "31 Mar 2025")
#                         date_str = parts[0].strip()
#                         description = parts[1].strip()
#                         amount = float(parts[2].strip())
#                         currency = parts[3].strip()
#                         transaction_type = parts[4].strip().lower()

#                         # Convert date to standard format
#                         try:
#                             date_obj = datetime.strptime(date_str, "%d %b %Y")
#                             # Check if the date exists
#                             if date_obj.day != int(date_str.split()[0]):
#                                 print(
#                                     f"Warning: Invalid date"
#                                     f"'{date_str}' - skipping"
#                                 )
#                                 continue
#                         except ValueError:
#                             print(
#                                 f"Warning: Error parsing date '{date_str}'"
#                                 f"- skipping"
#                             )
#                             continue

#                         date_formatted = date_obj.strftime("%Y-%m-%d")

#                         # Categorize
#                         category = categorize(description)

#                         transaction = {
#                             "date": date_formatted,
#                             "desc": description[:30],
#                             "amount": amount,
#                             "type": (
#                                 "income"
#                                 if transaction_type == "credit"
#                                 else "expense"
#                             ),
#                             "category": category,
#                         }

#                         transactions.append(transaction)

#                         # Track daily categories for expenses
#                         if transaction_type != "credit":
#                             daily = daily_categories[date_formatted]
#                             daily[category] += amount

#                 except (ValueError, IndexError) as e:
#                     print(f"Warning: Error parsing line '{line}' - {e}")
#                     continue

#     except Exception as e:
#         print(f"Error loading transactions: {e}")
#         return [], defaultdict(lambda: defaultdict(float))

#     return transactions, daily_categories
# def load_transactions(file_path_or_object):
#     """Load transactions from uploaded file with proper CSV parsing"""
#     transactions = []
#     daily_categories = defaultdict(lambda: defaultdict(float))

#     try:
#         # Handle both file objects and file paths
#         if hasattr(file_path_or_object, "read"):
#             # File object - rewind to beginning and read content
#             file_path_or_object.seek(0)
            
#             # Read content in chunks for large files
#             content = b''
#             while True:
#                 chunk = file_path_or_object.read(8192)  # 8KB chunks
#                 if not chunk:
#                     break
#                 content += chunk
            
#             if isinstance(content, bytes):
#                 content = content.decode("utf-8")
#             lines = content.split("\n")
#         else:
#             # File path
#             with open(file_path_or_object, "r", encoding="utf-8") as file:
#                 lines = file.readlines()

#         # Parse CSV lines
#         for line in lines:
#             line = line.strip()
#             if line and not line.startswith(("#", "Date")):
#                 try:
#                     parts = line.split(",")
#                     if len(parts) >= 5:
#                         # Parse date (assuming format: "31 Mar 2025")
#                         date_str = parts[0].strip()
#                         description = parts[1].strip()
#                         amount = float(parts[2].strip())
#                         currency = parts[3].strip()
#                         transaction_type = parts[4].strip().lower()

#                         # Convert date to standard format
#                         try:
#                             date_obj = datetime.strptime(date_str, "%d %b %Y")
#                         except ValueError:
#                             print(f"Warning: Error parsing date '{date_str}' - skipping")
#                             continue

#                         date_formatted = date_obj.strftime("%Y-%m-%d")

#                         # Categorize
#                         category = categorize(description)

#                         transaction = {
#                             "date": date_formatted,
#                             "desc": description[:30],
#                             "amount": amount,
#                             "type": (
#                                 "income"
#                                 if transaction_type == "credit"
#                                 else "expense"
#                             ),
#                             "category": category,
#                         }

#                         transactions.append(transaction)

#                         # Track daily categories for expenses
#                         if transaction_type != "credit":
#                             daily = daily_categories[date_formatted]
#                             daily[category] += amount

#                 except (ValueError, IndexError) as e:
#                     print(f"Warning: Error parsing line '{line}' - {e}")
#                     continue

#     except Exception as e:
#         print(f"Error loading transactions: {e}")
#         return [], defaultdict(lambda: defaultdict(float))

#     return transactions, daily_categories

def load_transactions(file_path_or_object):
    """Load transactions from uploaded file with proper CSV parsing"""
    transactions = []
    daily_categories = defaultdict(lambda: defaultdict(float))

    try:
        # Handle both file objects and file paths
        if hasattr(file_path_or_object, "read"):
            # File object - rewind to beginning for mobile devices
            file_path_or_object.seek(0)
            
            # Read content in chunks for better memory handling
            content = b''
            while True:
                chunk = file_path_or_object.read(8192)  # 8KB chunks
                if not chunk:
                    break
                content += chunk
            
            if isinstance(content, bytes):
                try:
                    content = content.decode("utf-8")
                except UnicodeDecodeError:
                    # Try other encodings if UTF-8 fails
                    try:
                        content = content.decode("latin-1")
                    except:
                        print("❌ Error decoding file content")
                        return [], defaultdict(lambda: defaultdict(float))
            
            lines = content.splitlines()
        else:
            # File path
            try:
                with open(file_path_or_object, "r", encoding="utf-8") as file:
                    lines = file.readlines()
            except UnicodeDecodeError:
                # Try other encodings
                try:
                    with open(file_path_or_object, "r", encoding="latin-1") as file:
                        lines = file.readlines()
                except Exception as e:
                    print(f"❌ Error reading file: {e}")
                    return [], defaultdict(lambda: defaultdict(float))

        # Parse CSV lines
        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            if line and not line.startswith(("#", "Date", "Date,")):
                try:
                    # More robust CSV parsing
                    parts = [part.strip() for part in line.split(',')]
                    if len(parts) < 5:
                        continue

                    # Parse date (assuming format: "31 Mar 2025")
                    date_str = parts[0]
                    description = parts[1]
                    
                    try:
                        amount = float(parts[2])
                    except ValueError:
                        continue
                    
                    currency = parts[3]
                    transaction_type = parts[4].lower()

                    # Convert date to standard format
                    try:
                        date_obj = datetime.strptime(date_str, "%d %b %Y")
                        date_formatted = date_obj.strftime("%Y-%m-%d")
                    except ValueError:
                        # Try other date formats if needed
                        continue

                    # Categorize
                    category = categorize(description)

                    transaction = {
                        "date": date_formatted,
                        "desc": description[:30],
                        "amount": amount,
                        "type": "income" if transaction_type == "credit" else "expense",
                        "category": category,
                    }

                    transactions.append(transaction)

                    # Track daily categories for expenses
                    if transaction_type != "credit":
                        daily = daily_categories[date_formatted]
                        daily[category] += amount

                except (ValueError, IndexError) as e:
                    print(f"⚠️ Warning: Error parsing line {line_num}: {e}")
                    continue

    except Exception as e:
        print(f"❌ Error loading transactions: {e}")
        return [], defaultdict(lambda: defaultdict(float))

    print(f"✅ Loaded {len(transactions)} transactions")
    return transactions, daily_categories

def get_operation_status(
    analysis_success, month_sheet_success, summary_sheet_success
):
    """Return operation status message"""
    if analysis_success and month_sheet_success and summary_sheet_success:
        return ("✅ Analysis completed successfully  "
                "and data written to Google Sheets")
    elif (
        analysis_success
        and month_sheet_success
        and not summary_sheet_success
    ):
        return ("⚠️ Analysis completed, Month sheet updated "
                "but failed to update Summary sheet")
    elif (
        analysis_success and not month_sheet_success and summary_sheet_success
    ):
        return ("⚠️ Analysis completed, Summary sheet updated "
                "but failed to update Month sheet")
    elif (
        analysis_success
        and not month_sheet_success
        and not summary_sheet_success
    ):
        return (
            "⚠️ Analysis completed but failed to write data to Google Sheets"
        )
    elif (
        not analysis_success and month_sheet_success and summary_sheet_success
    ):
        return ("⚠️ Analysis failed but Google Sheets operations completed")
    elif (
        not analysis_success
        and month_sheet_success
        and not summary_sheet_success
    ):
        return ("⚠️ Analysis failed, Month sheet updated "
                "but failed to update Summary sheet")
    elif (
        not analysis_success
        and not month_sheet_success
        and summary_sheet_success
    ):
        return ("⚠️ Analysis failed, Summary sheet updated "
                "but failed to update Month sheet")


def get_operation_status(
    analysis_success, month_sheet_success, summary_sheet_success
):
    """Return operation status message"""
    if analysis_success and month_sheet_success and summary_sheet_success:
        return (
            "✅ Analysis completed successfully  "
            "and data written to Google Sheets"
        )
    elif (
        analysis_success
        and month_sheet_success
        and not summary_sheet_success
    ):
        return ("⚠️ Analysis completed, Month sheet updated "
                "but failed to update Summary sheet")
    elif (
        analysis_success and not month_sheet_success and summary_sheet_success
    ):
        return ("⚠️ Analysis completed, Summary sheet updated "
                "but failed to update Month sheet")
    elif (
        analysis_success
        and not month_sheet_success
        and not summary_sheet_success
    ):
        return (
            "⚠️ Analysis completed but failed to write data to Google Sheets"
        )
    elif (
        not analysis_success and month_sheet_success and summary_sheet_success
    ):
        return ("⚠️ Analysis failed but Google Sheets operations completed")
    elif (
        not analysis_success
        and month_sheet_success
        and not summary_sheet_success
    ):
        return ("⚠️ Analysis failed, Month sheet updated "
                "but failed to update Summary sheet")
    elif (
        not analysis_success
        and not month_sheet_success
        and summary_sheet_success
    ):
        return ("⚠️ Analysis failed, Summary sheet updated "
                "but failed to update Month sheet")
    else:
        return "❌ All operations failed"


def run_full_analysis_with_file(month, file_path, temp_dir, operation_id):
    """Full processing in background mode using uploaded file"""
    global OPERATION_STATUS

    analysis_success = False
    month_sheet_success = False
    summary_sheet_success = False

    OPERATION_STATUS[operation_id] = (
        "⏳ Processing started... Google Sheets update in background"
    )

    try:
        print(
            f"🚀 Starting FULL background analysis "
            f"for {month} with uploaded file"
        )
        transactions, daily_categories = load_transactions(file_path)

        if not transactions:
            print("No transactions found in uploaded file")
            return analysis_success, month_sheet_success, summary_sheet_success

        data = analyze(transactions, daily_categories, month)
        analysis_success = True

        print(f"{month.upper()} ANALYSIS COMPLETED")
        print(f"Income: {data['income']:.2f}€")
        print(f"Expenses: {data['expenses']:.2f}€")
        print(f"Savings: {data['savings']:.2f}€")

        # 1. Writing into month sheet
        print(f"📝 Writing to {month} worksheet...")
        month_sheet_success = write_to_month_sheet(month, transactions, data)
        if month_sheet_success:
            OPERATION_STATUS[operation_id] = (
                "⏳ Month sheet updated, updating Summary..."
            )
        else:
            print(f"❌ Failed to update {month} worksheet")
            OPERATION_STATUS[operation_id] = (
                "❌ Failed to update Month worksheet"
            )

        time.sleep(10)

        # 2. Writing into Summary sheet
        print("⏳ Starting Google Sheets SUMMARY update...")
        table_data = prepare_summary_data(data, transactions)
        MONTH_NORMALIZED = get_month_column_name(month)
        summary_sheet_success = write_to_target_sheet(
            table_data, MONTH_NORMALIZED
        )

        if summary_sheet_success:

            OPERATION_STATUS[operation_id] = (
                "✅ Google Sheets update completed successfully!"
            )
        else:
            print("❌ Failed to update Google Sheets SUMMARY")
            OPERATION_STATUS[operation_id] = (
                "❌ Failed to update Google Sheets"
            )

        # Printing status message
        status_message = get_operation_status(
            analysis_success, month_sheet_success, summary_sheet_success
        )
        print(f"🎉 {status_message}")

    except Exception as e:
        print(f"Background analysis error: {e}")
        OPERATION_STATUS[operation_id] = (
            f"❌ Error during processing: {str(e)}"
        )
        import traceback

        print(f"Traceback: {traceback.format_exc()}")

    finally:
        # Clearing temporary data
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                print(f"Cleaned up temporary directory: {temp_dir}")
        except Exception as cleanup_error:
            print(f"Error cleaning up temporary files: {cleanup_error}")

    return analysis_success, month_sheet_success, summary_sheet_success


HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Finance Analyzer</title>
    <style>
        html {
                scroll-behavior: smooth;
            }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .main-container {
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
            overflow: hidden;
            width: 700px;
            max-width: 95%;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            padding: 25px;
            text-align: center;
            color: white;
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 600;
            letter-spacing: 1px;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
        }
        .header p {
            margin: 10px 0 0 0;
            font-size: 14px;
            opacity: 0.9;
        }
        .content {
            padding: 30px;
        }
        .form-container {
            text-align: center;
            margin-bottom: 25px;
        }
        .input-group {
            display: flex;
            flex-direction: column;
            gap: 15px;
            justify-content: center;
            align-items: center;
            margin-bottom: 20px;
        }
        input[type="text"], input[type="file"] {
            padding: 14px 20px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
            width: 300px;
            transition: all 0.3s ease;
            background: #f8f9fa;
        }
        input[type="text"]:focus, input[type="file"]:focus {
            outline: none;
            border-color: #667eea;
            background: white;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        input[type="file"] {
            padding: 12px;
            cursor: pointer;
        }
        button {
            padding: 14px 30px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }
        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
        }
        .terminal {
            background: #2d3748;
            color: #e2e8f0;
            border: 2px solid #667eea;
            padding: 10px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            line-height: 1.4;
            overflow: auto;
            white-space: pre-wrap;
            max-height: 700px;
            margin-bottom: 20px;
            scroll-margin-top: 20px;
            transition: all 0.3s ease;
        }
        @media (max-width: 768px) {
        .terminal {
            max-height: 90vh; /* 80% высоты экрана */
            height: auto;
            min-height: 400px; /* Минимальная высота */
            font-size: 12px; /* Чуть меньший шрифт */
            line-height: 1.3;
            padding: 8px;

            width: 100vw !important;
            max-width: 100vw !important;
            margin-left: -20px;
            margin-right: -20px;
            border-radius: 0;
            border-left: none;
            border-right: none;
            
            font-size: 12px;
            line-height: 1.3;
        }
        
        /* Дополнительно: уменьшите отступы на мобильных */
        .content {
            padding: 15px;
        }
        
         /* Чтобы контейнер не обрезал контент */
        .main-container {
            overflow-x: hidden;
        }
        
        .content {
            padding-left: 0;
            padding-right: 0;
        }
    }

        .main-container {
            margin: 10px;
            width: 100%;
        }
    }

    /* Для очень маленьких экранов */
    @media (max-width: 480px) {
        .terminal {
            max-height: 90vh;
            min-height: 350px;
            font-size: 11px;

            
            padding: 10px;
            margin-left: -15px;
            margin-right: -15px;
        }
        
        input[type="text"], input[type="file"] {
            font-size: 14px; /* Увеличим для удобства касания */
        }
    }

    /* Для горизонтальной ориентации */
    @media (max-width: 768px) and (orientation: landscape) {
        .terminal {
            
            height: auto;
            font-size: 11px;
        }
    }

        .terminal:empty {
            display: none !important;
        }
        .status {
            text-align: center;
            padding: 15px;
            border-radius: 25px;
            font-weight: 500;
            margin: 10px 0;
            font-size: 16px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
            transition: all 0.3s ease;
        }
        .hidden { display: none !important; }
        .status-loading {
            background: #fff3cd;
            color: #856404;
            border: 2px solid #ffeaa7;
        }
        .status-success {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            color: #155724;
            border: 2px solid #b8d8be;
            box-shadow: 0 4px 15px rgba(76, 175, 80, 0.3);
        }
        .status-warning {
            background: #fff3cd;
            color: #856404;
            border: 2px solid #ffeaa7;
        }
        .status-error {
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
            color: #721c24;
            border: 2px solid #f1b0b7;
            box-shadow: 0 4px 15px rgba(244, 67, 54, 0.3);
        }
        .file-info {
            margin-top: 10px;
            padding: 10px;
            background: #e8f4f8;
            border-radius: 6px;
            border-left: 4px solid #2196F3;
        }
        .status-container {
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 15px 0;
        }
        .anchor {
            scroll-margin-top: 20px;
        }
    </style>
</head>
<body>
    <div class="main-container">
        <div class="header">
            <h1>💰 PERSONAL FINANCE ANALYZER</h1>
            <p>Upload your CSV file and analyze your finances</p>
        </div>
        <div class="content">
            <div class="form-container">
                <form method="POST"
                enctype="multipart/form-data"
                id="uploadForm">
                
                    <div class="input-group">
                        <input type="text" name="month"
                        placeholder="Enter month (e.g. March, April)" required>
                        <input type="file" name="file" accept=".csv" required>
                        <button type="submit" id="submitBtn">Analyze</button>
                    </div>
                </form>
                <div id="mobileRetry" style="display: none; text-align: center; margin: 20px;">
    <button onclick="window.location.reload()" 
            style="padding: 12px 24px; background: #667eea; color: white; border: none; border-radius: 8px;">
        🔄 Retry Upload
    </button>
</div>
                {% if filename %}
                <div class="file-info anchor" id="fileInfoSection">
                    📁 Using file: <strong>{{ filename }}</strong>
                </div>
                {% endif %}
            </div>

            {% if result %}
            <div class="terminal" id="resultsSection">
                {{ result|safe }}
            </div>
            {% endif %}

            <div class="status-container">
                <div class="status hidden" id="statusMessage">
                    Processing your financial data...
                    Google Sheets update in progress
                </div>
            </div>
        </div>
    </div>

    <script>
document.getElementById('uploadForm').addEventListener('submit', function(e) {
    const statusElement = document.getElementById('statusMessage');
    const submitBtn = document.getElementById('submitBtn');
    const terminalElement = document.querySelector('.terminal');
    const fileInput = document.querySelector('input[type="file"]');
    const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);

if (isMobile) {
    // Увеличиваем таймаут отправки формы для мобильных
    document.getElementById('uploadForm').addEventListener('submit', function(e) {
        setTimeout(function() {
            // Продолжаем обработку
        }, 1000);
    });
}

    if (terminalElement) {
        terminalElement.innerHTML = '';
        terminalElement.style.display = 'none';
    }

    if (fileInput.files.length > 0) {
        const fileName = fileInput.files[0].name;
        let fileInfoElement = document.querySelector('.file-info');

        if (!fileInfoElement) {
            fileInfoElement = document.createElement('div');
            fileInfoElement.className = 'file-info';
            document.querySelector('.input-group').after(fileInfoElement);
        }

        fileInfoElement.innerHTML = `📁 Using : <strong>${fileName}</strong>`;
        fileInfoElement.style.display = 'block';
    }

    statusElement.classList.remove('hidden');
    statusElement.classList.remove('status-success','status-error','status-warning');
    statusElement.classList.add('status-loading');
    statusElement.textContent = '⏳ Google Sheets update in progress';

    submitBtn.disabled = true;
    submitBtn.textContent = 'Processing...';
    submitBtn.style.opacity = '0.7';
});

{% if status_message %}
document.addEventListener('DOMContentLoaded', function() {
    const statusElement = document.getElementById('statusMessage');
    statusElement.classList.remove('hidden');
    statusElement.textContent = '{{ status_message }}';

    {% if 'success' in status_message %}
    statusElement.classList.add('status-success');
    {% elif 'failed' in status_message %}
    statusElement.classList.add('status-error');
    {% elif 'warning' in status_message %}
    statusElement.classList.add('status-warning');
    {% else %}
    statusElement.classList.add('status-loading');
    {% endif %}
});
{% endif %}

{% if operation_id %}

function scrollToFileInfo() {
    const fileInfoSection = document.getElementById('fileInfoSection');
    if (fileInfoSection) {
        setTimeout(() => {
            fileInfoSection.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }, 300);
    }
}
document.addEventListener('DOMContentLoaded', function() {
    {% if result %}
    scrollToFileInfo();
    {% endif %}
});
// Function to check the status of the operation
function checkOperationStatus(operationId) {
    fetch('/status/' + operationId)
        .then(response => response.json())
        .then(data => {
            const statusElement = document.getElementById('statusMessage');
            if (statusElement) {
                statusElement.textContent = data.status;

                // Updating classes based on status
                if (data.status.includes('✅')) {
                    statusElement.className = 'status status-success';
                } else if (data.status.includes('❌')) {
                    statusElement.className = 'status status-error';
                } else if (data.status.includes('⏳')) {
                    statusElement.className = 'status status-loading';
                    // We continue to check the status every 5 seconds.
                    setTimeout(() => checkOperationStatus(operationId), 5000);
                } else if (data.status.includes('⚠️')) {
                    statusElement.className = 'status status-warning';
                }
            }
        })
        .catch(error => {
            console.error('Error checking status:', error);
        });
}

// We are initiating a status check when the page loads.
document.addEventListener('DOMContentLoaded', function() {
    checkOperationStatus('{{ operation_id }}');
});
{% endif %}
document.addEventListener('DOMContentLoaded', function() {
    const terminal = document.querySelector('.terminal');
    if (terminal && /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)) {
        // Вычисляем доступную высоту для терминала
        const viewportHeight = window.innerHeight;
        const terminalTop = terminal.getBoundingClientRect().top;
        const availableHeight = viewportHeight - terminalTop - 30; // 30px отступ снизу
        
        terminal.style.maxHeight = availableHeight + 'px';
        terminal.style.fontSize = '12px';
    }
});
</script>

</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    month = None
    filename = None
    status_message = None
    operation_id = None

    try:
        if request.method == "POST":
            print("📨 POST request received")

            
            # Mobile device detection and delay
            user_agent = request.headers.get('User-Agent', '').lower()
            if 'android' in user_agent or 'mobile' in user_agent:
                print("📱 Mobile device detected - adding delay")
                time.sleep(3)  # 3 second delay for mobile devices
            
            month = request.form.get("month", "").strip().lower()
            
            if not month:
                return render_template_string(
                    HTML, result="Month is required", status_message="❌ Please enter a month"
                )

            print(f"📅 Month: {month}")

            if "file" not in request.files:
                return render_template_string(
                    HTML, result="No file uploaded", status_message="❌ No file selected"
                )

            file = request.files["file"]
            if file.filename == "":
                return render_template_string(
                    HTML, result="No file selected", status_message="❌ Please select a file"
                )

            # Enhanced file validation
            if not file or not allowed_file(file.filename):
                return render_template_string(
                    HTML, 
                    result="Invalid file type. Please upload a CSV file.", 
                    status_message="❌ Invalid file type"
                )
            

            # Check file size (max 10MB)
            file.seek(0, 2)  # Seek to end to get size
            file_size = file.tell()
            file.seek(0)  # Reset to beginning
            
            if file_size > 10 * 1024 * 1024:  # 10MB limit
                return render_template_string(
                    HTML,
                    result="File too large. Maximum size is 10MB.",
                    status_message="❌ File too large"
                )
            print(f"🔍 Checking file: {file.filename}")
            print(f"📏 File size: {file_size} bytes")
            print(f"📄 Content type: {file.content_type}")
            print(f"✅ Allowed check: {allowed_file(file.filename)}")

            if file and (allowed_file(file.filename) or 
                        file.filename.lower().endswith('.csv') or 
                        file.content_type in ['text/csv', 'application/vnd.ms-excel', 'text/plain']):
                print("✅ File accepted for processing")
            if file and allowed_file(file.filename):
                # ← ВСТАВЬТЕ ЗДЕСЬ
                print(f"📁 File received: {file.filename}")
                print(f"📏 File size: {file_size} bytes")
                print(f"🔍 File content type: {file.content_type}")


                
                try:
                    filename = secure_filename(file.filename)
                    # Create temporary file for processing
                    temp_dir = tempfile.mkdtemp()
                    temp_file_path = os.path.join(
                        temp_dir, f"hsbc_{month}.csv"
                    )
                    # Save uploaded file
                    file.save(temp_file_path)

                    # Load transactions for immediate display
                    transactions, daily_categories = load_transactions(
                        temp_file_path
                    )

                    if transactions:
                        data = analyze(transactions, daily_categories, month)
                        result = format_terminal_output(
                            data, month, len(transactions)
                        )

                        # Generate unique operation ID
                        operation_id = f"{month}_{
                            datetime.now().strftime('%Y%m%d_%H%M%S')}"

                        # Start background processing
                        thread = threading.Thread(
                            target=run_full_analysis_with_file,
                            args=(
                                month,
                                temp_file_path,
                                temp_dir,
                                operation_id,
                            ),
                        )
                        thread.daemon = True
                        thread.start()

                        status_message = f"⏳ Processing started... "
                        f"Operation ID: {operation_id}"
                    else:
                        result = f"No valid transactions found in {filename}"
                        status_message = (
                            "❌ Analysis failed - no transactions found"
                        )

                except Exception as e:
                    result = f"Error processing file: {str(e)}"
                    status_message = "❌ Analysis failed due to error"
            else:
                result = "Invalid file type. Please upload a CSV file."
                status_message = "❌ Invalid file type"

        return render_template_string(
            HTML,
            result=result,
            month=month,
            filename=filename,
            status_message=status_message,
            operation_id=operation_id,
            
        )

    except Exception as e:
        print(f"Error in index function: {e}")
        return render_template_string(
            HTML,
            result=f"Error: {str(e)}",
            status_message="❌ System error occurred",
        )


@app.route("/status/<operation_id>")
def check_status(operation_id):
    """Check the status of a background operation"""
    global OPERATION_STATUS
    status = OPERATION_STATUS.get(operation_id, "Operation not found")
    return {"status": status}


def main():
    if "DYNO" in os.environ:
        # Heroku mode
        port = int(os.environ.get("PORT", 5000))
        app.run(host="0.0.0.0", port=port)
    else:
        # Local mode
        print(f"PERSONAL FINANCE ANALYZER")
        MONTH = (
            input("Enter the month (e.g. 'March, April, May'): ")
            .strip()
            .lower()
        )
        FILE = f"hsbc_{MONTH}.csv"
        print(f"Loading file: {FILE}")

        transactions, daily_categories = load_transactions(FILE)
        if not transactions:
            print(f"No transactions found")
            sys.exit(1)

        data = analyze(transactions, daily_categories, MONTH)
        terminal_visualization(data)

        print("=" * 50)

        # 1. Writing into month sheet
        monthly_success = write_to_month_sheet(MONTH, transactions, data)
        if monthly_success:
            print(f"✅ Successfully updated {MONTH} worksheet")
        else:
            print(f"❌ Failed to update {MONTH} worksheet")

        time.sleep(10)

        # 2. Writing into Summary sheet
        print("⏳ Writing to Google Sheets SUMMARY...")
        table_data = prepare_summary_data(data, transactions)
        MONTH_NORMALIZED = get_month_column_name(MONTH)

        success = write_to_target_sheet(table_data, MONTH_NORMALIZED)

        if success:
            print("✅ Google Sheets update completed successfully!")
        else:
            print("❌ Failed to update Google Sheets")


if __name__ == "__main__":
    if "DYNO" in os.environ:
        # Heroku mode
        port = int(os.environ.get("PORT", 5000))
        app.run(host="0.0.0.0", port=port)
    else:
        # Local mode
        main()
