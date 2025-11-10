# Import all requirements
from datetime import datetime, timedelta
import os
import pandas as pd
from slack_scripts import slack_api_run
from utilse import slack_open_file, open_file, get_first_last_dates, search_file, generate_excel

# Update according to your system
default_directory = "C:\\Users\\bridg\\project\\Salary-Generator\\Salary_data\\"
downloads_directory = "C:\\Users\\bridg\\Downloads\\"

# It should be updated annually to reflect any changes or new holidays for the current year.
# Make sure to review and modify this list at the start of each year to ensure accurate attendance management.
holidays = ['2025-01-01','2025-03-14','2025-03-19','2025-08-15','2025-10-02','2025-10-20','2025-10-22','2025-10-23','2025-12-25'] # add holidays in this format "YYYY-MM-DD" (Example "2025-01-01")


def generate_monthly_records(file_name):
    attendance_data = {}
    final_data = []

    # Path to the Excel file (replace this with the actual file path)
    excel_path = os.path.join(default_directory, f"{file_name}_DAILY_REPORT.xlsx")

    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_path, engine='openpyxl')

        # Check if the required columns exist in the Excel file
        if set(['Employee_ID', 'Status', 'Duration', 'Name']).issubset(df.columns):
            # Iterate through each row in the DataFrame
            for _, row in df.iterrows():
                emp_id = row['Employee_ID']
                status = row['Status']
                duration_str = row['Duration']
                name = row['Name']

                # Convert duration to timedelta
                try:
                    hours, minutes = map(int, duration_str.split(':'))
                    duration = timedelta(hours=hours, minutes=minutes)
                except ValueError:
                    print(f"Invalid duration format for employee {emp_id}: {duration_str}")
                    continue  # Skip the row if duration format is incorrect

                # Initialize employee data if not already present
                if emp_id not in attendance_data:
                    attendance_data[emp_id] = {
                        'name': name,
                        'total_duration': timedelta(),
                        'total_full_days': 0,
                        'total_half_days': 0,
                        'total_absent': 0,
                        'total_work_days': 0,
                        'total_month_working_days': 0
                    }

                # Update total duration
                attendance_data[emp_id]['total_duration'] += duration

                # Count full day, half day, or absent
                if status == 'P' or status == "Sat/Sun + Holiday" or status == "Sat/Sun" or status == "Holiday":
                    attendance_data[emp_id]['total_full_days'] += 1
                    attendance_data[emp_id]['total_work_days'] += 1
                elif status == 'HD':
                    attendance_data[emp_id]['total_half_days'] += 1
                    attendance_data[emp_id]['total_work_days'] += 0.5
                else:
                    attendance_data[emp_id]['total_absent'] += 1

                attendance_data[emp_id]['total_month_working_days'] += 1

            # Add header row to final data
            final_data.append([
                "S.No.",
                "ID",
                "Name",
                "Total Work Hours",
                "Full Days",
                "Half Days",
                "Absent",
                "Total Working Days(FD+HD)",
                "Total Month Working Days"
            ])

            # Process each employee's data
            sr_no = 1
            for emp_id, data in attendance_data.items():
                total_duration = data['total_duration']
                total_hours = f"{int(total_duration.total_seconds() // 3600)}:{(int(total_duration.seconds % 3600) // 60)}"
                final_data.append([
                    sr_no,
                    emp_id,
                    data['name'],
                    total_hours,
                    data['total_full_days'],
                    data['total_half_days'],
                    data['total_absent'],
                    data['total_work_days'],
                    data['total_month_working_days']
                ])
                sr_no += 1

            # Define the output file path
            excel_file_path = f"{file_name}_MONTHLY_REPORT.xlsx"

            # Write to Excel
            df_final = pd.DataFrame(final_data)
            df_final.to_excel(excel_file_path, index=False, header=False)
            print(f"Excel File generated as {excel_file_path}")

        else:
            print("Required columns not found in the Excel file.")

    except FileNotFoundError:
        print("The Excel file was not found. Please make sure the path is correct.")
    except Exception as e:
        print(f"An error occurred: {e}")



# Function to find file and process. If file not found the rest part will be skipped and it will  return file not found. 
def main_function():

    print('\nFile name must like : Example --> "JANUARY_2024" or "JAN_2024"')
    filename = input("Enter your file name (file extension optional): ")

    # Determine file extension and construct full paths
    if "." not in filename:
        # If the filename doesn't include an extension, assume .txt
        file_path = os.path.join(default_directory, f"{filename}.txt")
        file_path_downloads = os.path.join(downloads_directory, f"{filename}.txt")
        filename += ".txt"
    else:
        # If the filename includes an extension, use it as is
        file_path = os.path.join(default_directory, filename)
        file_path_downloads = os.path.join(downloads_directory, filename)


    # Attempt to open and read the file
    try:
        # Try to open the file from the default directory
        if os.path.exists(file_path):
            lines = open_file(file_path)
                
        # If not found in the default directory, try the downloads folder
        elif os.path.exists(file_path_downloads):
            lines = open_file(file_path_downloads)

        # If not found in either directory, search the entire system
        else:
            print('\nFile not found in Salary_data and Downloads folders.\nNow searching in C:\\\\ drive of this system... please wait!!')
            file_location = search_file("C:\\", filename)  # Start searching from C:\\
            if file_location:
                lines = open_file(file_location)  
            else:
                return "\nFile not found!!"
    
        # Get user input for start_date and end_date
        file_name = filename.split('.')[0]
        start_dt, end_dt = get_first_last_dates(file_name)

        start_date = datetime.strptime(start_dt, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_dt, "%Y-%m-%d").date()

        # call slack api to generate slack messages
        slack_api_run(file_name, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))

        print(f"\nProcessing {file_name} from {start_date} to {end_date}")

        month_name = filename.split('_')[0]
        # To get month nummber
        try:
            # Try full month name
            month_number = datetime.strptime(month_name.capitalize(), "%B").month
        except ValueError:
            # If it fails, try short month name
            month_number = datetime.strptime(month_name.upper(), "%b").month
        month_holidays = []
        for date in holidays:
            if int(date[5:7]) == month_number:
                month_holidays.append(date)

        # Showing Holidays of this current month to the user 
        print(f"Holidays in {file_name} : {month_holidays}\n")

        # Generate the list of weekdays
        sat_sun = []
        date_list_ex_sat_sun = []
        date_list_include_sat_sun = []
        current_date = start_date
        while current_date <= end_date:
            if current_date.weekday() < 5:  # 0 = Monday, 4 = Friday
                date_list_ex_sat_sun.append(current_date)
                date_list_include_sat_sun.append(current_date)
            else:
                sat_sun.append(current_date)
                date_list_include_sat_sun.append(current_date)
            current_date += timedelta(days=1)

        # Remove closed dates from the date_list
        # date_list = [date for date in date_list_include_sat_sun if str(date) not in holidays]

        # Create all required variables here
        employee_working_hours = {}
        employee_names = {}
        final = {}

        # Set Half Day and Full Day thresholds
        full_day_threshold = timedelta(hours=7, minutes=45) # Full Day working hours
        half_day_threshold = timedelta(hours=4) # Half Day working hours

        # Separate Important parts
        for line in lines[1:]:
            parts = line.split("\t")
            emp_id = parts[2]
            name = parts[3].title()
            date = datetime.strptime(parts[-1].strip(), "%Y-%m-%d %H:%M:%S").date()

            # Track employee names
            if emp_id not in employee_names:
                employee_names[emp_id] = name

            # Process only if date is in our list
            if date in date_list_include_sat_sun:
                if emp_id not in employee_working_hours:
                    employee_working_hours[emp_id] = {}
                if date not in employee_working_hours[emp_id]:
                    employee_working_hours[emp_id][date] = {"totalEntry": [], "name": name}
                time = datetime.strptime(parts[-1].strip(), "%Y-%m-%d %H:%M:%S").time()
                entry_key = f"{time.strftime('%H:%M:%S')}"
                employee_working_hours[emp_id][date]["totalEntry"].append(entry_key)

        # Create final dictionary ensuring all weekdays are included
        for emp_id in employee_names:
            for date in date_list_include_sat_sun:
                if date not in employee_working_hours.get(emp_id, {}):
                    status = 'A'
                    if date in sat_sun and str(date) in month_holidays:
                        status = "Sat/Sun + Holiday"
                    elif date in sat_sun:
                        status = "Sat/Sun"
                    elif str(date) in month_holidays:
                        status = "Holiday"
                    else:
                        pass

                    # Initialize absent day entry
                    final.setdefault(emp_id, {})[date] = {
                        'name': employee_names[emp_id],
                        'duration': '00:00',
                        'total punch': [],
                        'punch': 'none',
                        'Status': status
                    }
                else:
                    data = employee_working_hours[emp_id][date]
                    length_of_totalEntry = len(data['totalEntry'])
                    total_duration = timedelta()

                    # Odd number of entries: skip the last entry for duration calculation
                    if length_of_totalEntry % 2 != 0:
                        total_entries_for_duration = data['totalEntry'][:-1]
                        if len(total_entries_for_duration) > 1:
                            for i in range(0, len(total_entries_for_duration), 2):
                                odd_entry_time = datetime.strptime(total_entries_for_duration[i], "%H:%M:%S").time()
                                even_entry_time = datetime.strptime(total_entries_for_duration[i + 1], "%H:%M:%S").time()
                                odd_time = datetime.combine(datetime.today(), odd_entry_time)
                                even_time = datetime.combine(datetime.today(), even_entry_time)
                                time_difference = even_time - odd_time
                                total_duration += time_difference
                    else:
                        odd_entry_time = []
                        even_entry_time = []
                        for i in range(0, length_of_totalEntry, 2):
                            odd_entry_time.append(datetime.strptime(data['totalEntry'][i], "%H:%M:%S").time())
                            even_entry_time.append(datetime.strptime(data['totalEntry'][i + 1], "%H:%M:%S").time())
                        for e, o in zip(even_entry_time, odd_entry_time):
                            even_time = datetime.combine(datetime.today(), e)
                            odd_time = datetime.combine(datetime.today(), o)
                            time_difference = even_time - odd_time
                            total_duration += time_difference

                    # Calculating duration
                    total_seconds = total_duration.total_seconds()
                    total_hours, remainder = divmod(total_seconds, 3600)
                    total_minutes = remainder // 60
                    duration_str = "{:02}:{:02}".format(int(total_hours), int(total_minutes))

                    # Determine Full Day, Half Day, or Absent according to calculated duration
                    if total_duration >= full_day_threshold:
                        status = 'P'
                    elif total_duration >= half_day_threshold:
                        status = 'HD'
                    else:
                        status = 'A'
                    
                    if date in sat_sun and str(date) in month_holidays:
                        status = "Sat/Sun + Holiday"
                    elif date in sat_sun:
                        status = "Sat/Sun"
                    elif str(date) in month_holidays:
                        status = "Holiday"
                    else:
                        pass

                    # Adding data in the final dictionary
                    final.setdefault(emp_id, {})[date] = {
                        'name': data['name'],
                        'Status': status,
                        'duration': duration_str,
                        'total punch': employee_working_hours[emp_id][date]["totalEntry"],
                        'punch': 'Odd' if length_of_totalEntry % 2 != 0 else 'Even',
                    }

        slack_path = os.path.join(default_directory, f"{file_name}_slack_messages.csv")
        # Export this data to CSV
        data_list = []
        for emp_id, dates in final.items():
            for date, data in dates.items():
                # issue = []
                issue = ""
                slack_messages = slack_open_file(slack_path)
                for row in slack_messages:
                    row=row.split(',')
                    if row[2] == str(date) and row[1].title() == data["name"].title():
                        # issue.append(f"Time: {row[3]} Message: {row[4]}")
                        issue += f"Time: {row[3]} Message: {row[4]} "
                entry = {
                    'Employee_ID': emp_id,
                    'Name': data['name'],
                    'Date': date.strftime('%Y-%m-%d'),  # Format date as YYYY-MM-DD
                    'Status': data.get('Status'),
                    'Duration': data['duration'],
                    "Total_Punch": data['total punch'],
                    'Punch_type': data['punch'],
                    "Punch_issues": issue
                }

                data_list.append(entry)

        # Sorting the order of entries in the final data
        sorted_data_list = sorted(data_list, key=lambda x: (x['Name'], x['Date']))

        # Generating file
        generate_excel(file_name, sorted_data_list)

        generate_monthly_records(file_name)

    except FileNotFoundError:
        print('File not found in any of the directories in C\\: drive of this system.')
    except UnicodeDecodeError:
        print('Error reading the file. It might not be encoded in UTF-16. Try a different encoding.')
    except Exception as e:
        print(f"An error occurred: {e}")
    print("\nProcess Completed!!")
   
main_function()