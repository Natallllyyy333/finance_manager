# Write your code to expect a terminal of 80 characters wide and 24 rows high
import csv
MONTH = input(
    "Enter the month (e.g., 'March', 'April', 'May'): ").strip().lower()

file = f"hsbc_{MONTH}.csv"

with open(file, mode='r') as csv_file:
    csv_reader = csv.reader(csv_file)
    for row in csv_reader:
        print(row)
