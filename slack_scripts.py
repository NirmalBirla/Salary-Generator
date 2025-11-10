import csv
from datetime import datetime
import os
import pytz
import re

from slack_api import fetch_all_users, retrieve_messages

def convert_to_kolkata_time(timestamp_in_seconds):
    # Convert the timestamp into a UTC datetime object
    utc_time = datetime.utcfromtimestamp(float(timestamp_in_seconds))

    # Define the Asia/Kolkata timezone
    kolkata_timezone = pytz.timezone('Asia/Kolkata')

    # Convert UTC time to Asia/Kolkata time
    kolkata_time = pytz.utc.localize(utc_time).astimezone(kolkata_timezone)

    # Extract date and time as separate objects
    date = kolkata_time.strftime('%Y-%m-%d')  
    time = kolkata_time.strftime('%H:%M:%S')
    
    return date, time

def messages_in_date_range(all_slack_messages, start_date, end_date):
    # Filter messages based on date range

    start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
    end_date = datetime.strptime(end_date, "%Y-%m-%d").date()

    filtered_messages = []
    for message in all_slack_messages:
        message_date, kolkata_timestamp = convert_to_kolkata_time(message["ts"])

        if str(start_date) <= message_date <= str(end_date):
            filtered_messages.append(message)

    return filtered_messages


def save_to_csv(filtered_messages, filename, users_list):
    # Save filtered Slack data to a CSV file
    try:
        with open(filename, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["UserId", "Full Name", "Date", "Time", "Message"])

            for message in filtered_messages:
                user_id = message.get("user")
                user_info = users_list.get(user_id, {"real_name": "Unknown", "email": "N/A"})
                full_name = user_info.get("real_name", "Unknown").strip()

                # Convert timestamp to Kolkata time
                date, time = convert_to_kolkata_time(message["ts"])

                # Clean the message text and replace user mentions
                new_message = message.get("text", "").strip()
                
                for uid, user_data in users_list.items():
                    mention = f"<@{uid}>"
                    if mention in new_message:
                        new_message = new_message.replace(mention, f"@{user_data['real_name']}")
                
                # Remove commas from the message to avoid CSV issues
                new_message = new_message.replace(",", "")

                # Remove any leading/trailing spaces and multiple spaces between words
                new_message = re.sub(r'\s+', ' ', new_message).strip()
                # Write the row to the CSV file
                writer.writerow([user_id, full_name, date, time, new_message])
        file_name = filename.split('\\')[-1]
        print(f"\nSlack Messages File generated as {file_name}")
    
    except Exception as e:
        print(f"An error occurred while saving to CSV: {e}")



def slack_api_run(file_name, start_date, end_date):
    try:
        # Fetch all users from Slack
        users_list = fetch_all_users()
        if not users_list:
            print("Error: No users were fetched from Slack.")
            return

        # Retrieve Slack messages
        all_slack_messages = retrieve_messages()
        if not all_slack_messages:
            print("Error: No Slack messages retrieved.")
            return

        # Filter messages by date range
        filtered_messages = messages_in_date_range(all_slack_messages, start_date, end_date)
        if not filtered_messages:
            print(f"No messages found between {start_date} and {end_date}.")
            return

        # Define the default directory and file name
        default_directory = "C:\\Users\\bridg\\project\\Salary-Generator\\Salary_data\\"
        filename = f"{file_name}_slack_messages.csv"

        # Create the full file path
        file_path = os.path.join(default_directory, filename)

        # Save the filtered messages to a CSV file using the custom function
        save_to_csv(filtered_messages, file_path, users_list)
    
    except Exception as e:
        print(f"An error occurred during the Slack API run: {e}")

# file_name = 'OCTOBER_2025'
# start_date = '2025-10-01'
# end_date = '2025-10-31'
# slack_api_run(file_name, start_date, end_date)