import csv
from collections import defaultdict
from datetime import datetime
from dateutil.parser import parse
import sys
import os

# Function to read the CSV file and count calls per user per day by week and date
def count_calls_by_week_and_date(csv_file, selected_users=None, exempted_users=None):
    calls_by_week_and_date = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    if selected_users is None:
        selected_users = []
    if exempted_users is None:
        exempted_users = []

    with open(csv_file, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            upn = row['UPN']
            if not selected_users or upn in selected_users:
                if upn not in exempted_users:
                    start_time = parse(row['Start Time'])
                    week_number = start_time.strftime('%U')
                    day_of_week = start_time.strftime('%A')
                    date = start_time.date()
                    calls_by_week_and_date[upn][week_number][(day_of_week, date)] += 1

    return calls_by_week_and_date

# Function to display the report
def display_report(calls_by_week_and_date):
    print("Report of calls made by each user by week and date:")
    for upn, calls_by_week in calls_by_week_and_date.items():
        print(f"\nUser: {upn}")
        for week_number, calls in calls_by_week.items():
            print(f"\nWeek {week_number}")
            for (day_of_week, date), count in calls.items():
                date_str = date.strftime('%m-%d')
                print(f"{day_of_week} ({date_str}): {count} calls")


if __name__ == "__main__":
    if len(sys.argv) == 1:
        csv_files = [f for f in os.listdir() if f.endswith('.csv')]
        if len(csv_files) == 0:
            print("No CSV files found in the current directory")
            sys.exit(1)
        elif len(csv_files) == 1:
            csv_file = csv_files[0]
        else:
            print("Multiple CSV files found in the current directory:")
            for i, f in enumerate(csv_files):
                print(f"{i+1}. {f}")
            csv_number = int(input("Enter the number of the CSV file to use: "))
            if csv_number < 1 or csv_number > len(csv_files):
                print("Invalid input")
                sys.exit(1)
            csv_file = csv_files[csv_number - 1]
    else:
        csv_file = sys.argv[1]
    users = set()


    # Exempted User list, call groups, auto attendants, etc 
    exempted_users = [
    'user1@domain.com',
    'user2@domain.com'
    ]

    # Open CSV to get list to select from   
    with open(csv_file, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            users.add(row['UPN'])
    user_list = sorted(users - set(exempted_users))
    for i, user in enumerate(user_list):
        print(f"{i+1}. {user}")
    user_input = input("Enter comma-separated list of numbers (leave blank for all users): ")
    selected_numbers = [int(n.strip()) for n in user_input.split(',')] if user_input else None
    selected_users = [user_list[n-1] for n in selected_numbers] if selected_numbers else None
    calls_by_week_and_date = count_calls_by_week_and_date(csv_file, selected_users, exempted_users)
    display_report(calls_by_week_and_date)
