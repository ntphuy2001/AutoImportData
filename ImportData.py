import xlwings as xw
import pandas as pd
import os
import re
import json
import shutil
from typing import Dict, List, Any
from datetime import datetime, date
import calendar


# Get ticket id from an issue
def get_ticket_id(issue: str) -> str:
    """Extract ticket ID from issue description."""
    match = re.search(r'#\d+', issue)
    return match.group(0) if match else None


def init_members_data_in_month(members: Dict[str, str], month: int, year: int) -> Dict[str, List[List[Any]]]:
    """
    Initialize data structure for each member for a given month.

    Args:
    members (Dict[str, str]): A dictionary of member full names to nicknames.
    month (int): The month (1-12).
    year (int): The year.

    Returns:
    Dict[str, List[List[Any]]]: A dictionary with member nicknames as keys and lists of daily data as values.
    """
    # Get the number of days in the month
    _, days_in_month = calendar.monthrange(year, month)

    # Create a date object for the first day of the month
    first_day = date(year, month, 1)

    # Initialize the data structure using a dictionary comprehension
    return {
        nickname: [
            [
                first_day.replace(day=day),  # Date object for each day
                None,  # Code
                None,  # Start time
                None,  # End time
                []     # Tasks
            ]
            for day in range(1, days_in_month + 1)
        ]
        for nickname in members.values()
    }


def generateDataToEachMember(csv_file, members_data_in_month):
    # Read the CSV file
    df = pd.read_csv(csv_file).iloc[::-1]

    # Expected order of a row of data
    # templateData[0]: Date
    # templateData[1]: Code
    # templateData[2]: starttime
    # templateData[3]: endtime
    # templateData[4]: task

    for index, row in df.iterrows():
        if row['User'] not in members_data_in_month.keys():
            continue
        # Check if it is a tak of new day
        taskDate = datetime.strptime(row['Date'], '%m/%d/%Y')
        startDate = datetime.strptime('9:00', '%H:%M').time(),
        endDate = datetime.strptime('18:00', '%H:%M').time(),
        indexOfDayWorkTask = taskDate.day - 1
        members_data_in_month[row['User']][indexOfDayWorkTask][1] = 1
        members_data_in_month[row['User']][indexOfDayWorkTask][2] = startDate[0].strftime('%H:%M')
        members_data_in_month[row['User']][indexOfDayWorkTask][3] = endDate[0].strftime('%H:%M')
        members_data_in_month[row['User']][indexOfDayWorkTask][4].append(get_ticket_id(row['Issue']))

    return members_data_in_month


def copy_xlsm_file(app, source_path, destination_path):
    # Open the existing workbook
    wb = app.books.open(source_path)

    # Save the workbook with the new name
    wb.save(destination_path)

    # Close the workbook
    wb.close()

    # Optionally, you can ensure macros are copied by directly copying the file
    shutil.copyfile(source_path, destination_path)

    print(f"Copied {source_path} to {destination_path}")


def import_data(xlsm_file_path, csv_file_path):
    app = xw.App(visible=False)
    # List member

    config = open('config.json')
    data = json.load(config)
    members = data['members']

    xlsmFileName = os.path.splitext(xlsm_file_path)
    updated_xlsm_file_path = f"{xlsmFileName[0]}_update{xlsmFileName[1]}"

    copy_xlsm_file(app, xlsm_file_path, updated_xlsm_file_path)
    wb = app.books.open(updated_xlsm_file_path)

    sheet = wb.sheets['設定']
    year = int(sheet.range('C5').value)
    month = int(sheet.range('C7').value)

    membersDataInMonth = init_members_data_in_month(members, month, year)
    membersDataInMonth = generateDataToEachMember(csv_file_path, membersDataInMonth)

    # Access the sheet where you want to import data
    for fullname, nickname in members.items():
        sheet = wb.sheets[fullname]

        # Prepare data for writing in batches
        code = []
        starttime = []
        endtime = []
        task = []

        for day in membersDataInMonth[nickname]:
            code.append([day[1]])
            starttime.append([day[2]])
            endtime.append([day[3]])
            task.append([", ".join(day[4])] if day[4] != [] else [None])

        # Write data in batches
        sheet.range('D10').options(index=False).value = code
        sheet.range('F10').options(index=False).value = starttime
        sheet.range('G10').options(index=False).value = endtime
        sheet.range('K10').options(index=False).value = task

    # Save the Excel file
    wb.save()

    # Close the workbook without saving changes
    wb.close()

    # Quit the Excel application
    app.quit()