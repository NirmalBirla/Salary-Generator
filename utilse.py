from datetime import datetime, timedelta
import os
import csv
import pandas as pd

def slack_open_file(path):
    with open(path, "r", encoding="utf-8") as file:
        lines = file.readlines()
    return lines

# Function for open file with given path
def open_file(path):
    try:
        with open(path, "r", encoding="utf-8") as file:
            lines = file.readlines()
    except UnicodeDecodeError:
        with open(path, "r", encoding="utf-16") as file:
            lines = file.readlines()
    return lines

# Function to get first and last day of the month according to the given file name
def get_first_last_dates(file_name):
    # Split the input string into month and year
    month_name, year = file_name.split('_')
    year = int(year)
    
    # To get month nummber
    try:
        # Try full month name
        month_number = datetime.strptime(month_name.capitalize(), "%B").month
    except ValueError:
        # If it fails, try short month name
        month_number = datetime.strptime(month_name.upper(), "%b").month
    
    # Get the first date of the month
    first_date = datetime(year, month_number, 1)
    # first_date = datetime(2025, 10, 15, 0, 0) # To update the start date 
    
    # Get the last date of the month
    if month_number == 12:
        last_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        last_date = datetime(year, month_number + 1, 1) - timedelta(days=1)
    
    return first_date.strftime("%Y-%m-%d"), last_date.strftime("%Y-%m-%d")


# Function to search for a file in a directory and its subdirectories
def search_file(directory, filename):
    for root, dirs, files in os.walk(directory):
        if filename in files:
            return os.path.join(root, filename)
    return None

def generate_csv(file_name, sorted_data_list):
    # Generating a CSV file
    csv_file_path = f'{file_name}_DAILY_REPORT.csv'
    with open(csv_file_path, 'w', newline='') as csv_file:
        fieldnames = ['Employee_ID', 'Date', 'Status', 'Duration', 'Name', 'Total_Punch', 'Punch_type', 'Punch_issues']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for entry in sorted_data_list:
            writer.writerow(entry)
    print(f"CSV File generated as {csv_file_path}")

def generate_excel(file_name, sorted_data_list):
    df = pd.DataFrame(sorted_data_list)
    excel_file_path = f'{file_name}_DAILY_REPORT.xlsx'
    df.to_excel(excel_file_path, index=False)
    print(f"Excel File generated as {excel_file_path}")




import chardet

def find():
    # Detect encoding
    with open('AUGUST_2024.txt', 'rb') as f:
        raw_data = f.read()

    result = chardet.detect(raw_data)
    encoding = result['encoding']
    confidence = result['confidence']

    print(f"Detected encoding: {encoding} with confidence {confidence}")