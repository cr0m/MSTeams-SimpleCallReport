import os
import sys
import csv
import pytz
import win32com.client as win32
from datetime import datetime, timedelta
from dateutil.parser import parse
from collections import defaultdict

"""
Script Usage:

- '-inout': Number of outbound/inbound calls
  Example: python search.py CALLDATA.csv -inout

- '-talktime': Total talk time
  Example: python search.py CALLDATA.csv -talktime
"""

# Setup the Email Subject and Recipients
subject = "Weekly Call Report"
recipients = ["email@domain.com"]

# Exclude Users and Service Accounts
exempted_users = [
    'user1@domain.com',
    'user2@domain.com'
]

def create_outlook_email(subject, body, recipients=[]):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # 0 is for Mail item
    mail.Subject = subject
    mail.HTMLBody = body
    mail.To = "; ".join(recipients)  # Add recipients as needed
    mail.Display()  # Opens the email template

# Function to read the CSV file and count calls per user per day by week and date
def count_calls_by_week_and_date(csv_file, selected_users=None, exempted_users=None, start_date=None, end_date=None):
    calls_by_week_and_date = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {
        'count': 0, 'outbound': 0, 'inbound': 0, 'duration': 0, 'unique_numbers': set(), 'first_call': None, 'last_call': None
    })))

    if selected_users is None:
        selected_users = []
    if exempted_users is None:
        exempted_users = []

    eastern = pytz.timezone('US/Eastern')  # Define Eastern Timezone

    with open(csv_file, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            user_display_name = row['User Display Name']
            if not selected_users or user_display_name in selected_users:
                if user_display_name not in exempted_users:
                    # Parse and convert start time to Eastern Time
                    start_time_utc = parse(row['Start Time'])
                    start_time_est = start_time_utc.astimezone(eastern)

                    week_number = start_time_est.strftime('%U')
                    day_of_week = start_time_est.strftime('%A')
                    date = start_time_est.date()
                    if (start_date is None or date >= start_date) and (end_date is None or date <= end_date):
                        duration_seconds = int(row['Duration Seconds'])
                        phone_number = row['Destination Number']
                        call_direction = row['Call Direction']

                        calls_data = calls_by_week_and_date[user_display_name][week_number][(day_of_week, date)]
                        calls_data['count'] += 1
                        calls_data['duration'] += duration_seconds
                        calls_data['unique_numbers'].add(phone_number)

                        if call_direction == 'Outbound':
                            calls_data['outbound'] += 1
                        elif call_direction == 'Inbound':
                            calls_data['inbound'] += 1

                        # Update first and last call times (keep as datetime objects)
                        if calls_data['first_call'] is None or start_time_est < calls_data['first_call']:
                            calls_data['first_call'] = start_time_est
                        if calls_data['last_call'] is None or start_time_est > calls_data['last_call']:
                            calls_data['last_call'] = start_time_est

    return calls_by_week_and_date

def display_report(calls_by_week_and_date, show_duration=False, show_direction=False):
    bodyoutput = "" 
    print("Report of calls made by each user by week and date:")
    for user_display_name, calls_by_week in calls_by_week_and_date.items():
        for week_number, calls in sorted(calls_by_week.items()):
            total_outbound = 0
            total_inbound = 0
            dates_in_week = []

            for (day_of_week, date), stats in sorted(calls.items(), key=lambda x: x[0][1]):
                dates_in_week.append(date)
                total_outbound += stats['outbound']
                total_inbound += stats['inbound']

            if dates_in_week:
                week_start = min(dates_in_week)
                if week_start.weekday() != 6:
                    week_start = week_start - timedelta(days=week_start.weekday() + 1)
                week_end = week_start + timedelta(days=6)

            week_start_str = week_start.strftime('%m/%d')
            week_end_str = week_end.strftime('%m/%d')
            bodyoutput += f"<br><b><u>{user_display_name}</u></b> <br>"
            print(f"\n\n{user_display_name}") 
            bodyoutput += f"Week {week_number} ({week_start_str} - {week_end_str})<br>"
            print(f"Week {week_number} ({week_start_str} - {week_end_str})")
            bodyoutput += f"&nbsp;Total Outbound: {total_outbound} calls<br>"
            print(f"  Total Outbound: {total_outbound} calls")
            bodyoutput += f"&nbsp;Total Inbound: {total_inbound} calls<br>"
            print(f"  Total Inbound: {total_inbound} calls")

            for (day_of_week, date), stats in sorted(calls.items(), key=lambda x: x[0][1]):
                date_str = date.strftime('%m-%d')
                count = stats['count']
                outbound = stats['outbound']
                inbound = stats['inbound']
                unique_numbers_count = len(stats['unique_numbers'])
                total_calls = outbound + inbound
                percent_of_unique_numbers = (unique_numbers_count / total_calls * 100) if total_calls > 0 else 0
                duration = stats['duration']
                bodyoutput += f"<br>{day_of_week} ({date_str}):<br>"
                print(f"\n{day_of_week} ({date_str}):")
                bodyoutput += f"&nbsp;Outbound: {outbound} calls<br>"
                print(f"  Outbound: {outbound} calls")
                bodyoutput += f"&nbsp;&nbsp;Inbound: {inbound} calls<br>"
                print(f"  Inbound: {inbound} calls")
                bodyoutput += f"&nbsp;&nbsp;Unique Numbers: {unique_numbers_count} ({percent_of_unique_numbers:.2f}% unique)<br>"
                print(f"  Unique Numbers: {unique_numbers_count} ({percent_of_unique_numbers:.2f}% unique)")

                # Display first and last call times (formatted as strings)
                first_call_time = stats['first_call'].strftime('%I:%M:%S %p') if stats['first_call'] else 'N/A'
                last_call_time = stats['last_call'].strftime('%I:%M:%S %p') if stats['last_call'] else 'N/A'
                bodyoutput += f"&nbsp;&nbsp;First Call Time: {first_call_time}<br>"
                print(f"  First Call Time: {first_call_time}")
                bodyoutput += f"&nbsp;&nbsp;Last Call Time: {last_call_time}<br>"
                print(f"  Last Call Time: {last_call_time}")

                if show_duration:
                    duration_str = format_duration(duration)
                    bodyoutput += f"&nbsp;&nbsp;Total Talk Time: {duration_str}<br>"
                    print(f"  -- Total Talk Time: {duration_str}")

    create_outlook_email(subject, bodyoutput, recipients)

# Function to format duration from seconds to hours:minutes:seconds
def format_duration(seconds):
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours} hours, {minutes} minutes, {seconds} seconds"

# Main script
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

    with open(csv_file, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            users.add(row['User Display Name'])
        user_list = sorted(users - set(exempted_users))

    for i, user in enumerate(user_list):
        print(f"{i+1}. {user}")

    user_input = input("Enter comma-separated list of numbers (leave blank for all users): ")
    selected_numbers = [int(n.strip()) for n in user_input.split(',')] if user_input else None
    selected_users = [user_list[n-1] for n in selected_numbers] if selected_numbers else None

    report_type = input("Report on entire document or previous week only? [Enter 'all' for entire document or press Enter for previous week]: ").strip().lower()

    start_date = None
    end_date = None

    if report_type == '' or report_type == 'week':
        today = datetime.today().date()
        start_of_this_week = today - timedelta(days=today.weekday())
        start_date = start_of_this_week - timedelta(days=7)  # Sunday of the previous week
        end_date = today  # Current date
    elif report_type != 'all':
        print("Invalid input. Please enter 'all' for entire document or press Enter for previous week.")
        sys.exit(1)

    calls_by_week_and_date = count_calls_by_week_and_date(csv_file, selected_users, exempted_users, start_date, end_date)

    show_duration = "-talktime" in sys.argv
    show_direction = "-inout" in sys.argv

    report_output = display_report(calls_by_week_and_date, show_duration, show_direction)
