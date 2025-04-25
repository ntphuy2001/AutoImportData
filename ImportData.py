import numpy as np
import xlwings as xw
import pandas as pd
import os
import re
import json
import shutil
from typing import Dict, List, Any
from datetime import datetime, date, timedelta
import calendar
import logging


# Get ticket id from an issue
def get_ticket_id(issue: str) -> str:
    """Extract ticket ID from issue description."""
    match = re.search(r'#\d+', issue)
    return match.group(0) if match else None


def init_members_data_in_month(members: Dict[str, str], month: int, year: int, wb: xw.main.Books) \
        -> Dict[str, List[List[Any]]]:
    """
    Initialize data structure for each member for a given month.

    Args:
    members (Dict[str, str]): A List of member nicknames.
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
                []  # Tasks
            ]
            for day in range(1, days_in_month + 1)
        ]
        for nickname in members.values()
    }


def generate_data_to_each_member(logtimeData, members_data_in_month, date_format):
    try:
        # Expected order of a row of data
        # templateData[0]: Date
        # templateData[1]: Code
        # templateData[2]: starttime
        # templateData[3]: endtime
        # templateData[4]: task

        for index, row in logtimeData.iterrows():
            if row['User'] not in members_data_in_month.keys():
                continue
            # Check if it is a tak of new day
            taskDate = datetime.strptime(row['Date'], date_format)
            start_date = datetime.strptime('9:00', '%H:%M').time(),
            end_date = datetime.strptime('18:00', '%H:%M').time(),
            index_of_day_work_task = taskDate.day - 1
            members_data_in_month[row['User']][index_of_day_work_task][1] = 1
            if members_data_in_month[row['User']][index_of_day_work_task][2] is None:
                members_data_in_month[row['User']][index_of_day_work_task][2] = start_date[0].strftime('%H:%M')

            if members_data_in_month[row['User']][index_of_day_work_task][3] is None:
                start_datetime = datetime.combine(datetime.today(), start_date[0])
                task_duration = timedelta(hours=row['Hours'] + 1)
                end_datetime = start_datetime + task_duration
                members_data_in_month[row['User']][index_of_day_work_task][3] = end_datetime.strftime('%H:%M')
            else:
                start_datetime = datetime.combine(
                    datetime.today(),
                    datetime.strptime(members_data_in_month[row['User']][index_of_day_work_task][3], '%H:%M').time(),
                )
                task_duration = timedelta(hours=row['Hours'])
                end_datetime = start_datetime + task_duration
                members_data_in_month[row['User']][index_of_day_work_task][3] = end_datetime.strftime('%H:%M')
            members_data_in_month[row['User']][index_of_day_work_task][4].append(get_ticket_id(row['Issue']))

    except KeyError as e:
        logging.error(f"Invalid member key: {str(e)}")
    except ValueError as e:
        logging.error(f"Invalid data format: {str(e)}")
    except Exception as e:
        logging.error(f"An error occurred during data generation: {str(e)}")

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


def logtime_data(csv_file_path):
    return pd.read_csv(csv_file_path).iloc[::-1]


def write_data(sheet, code, starttime, endtime, tasks, start_row):
    for day in range(1, len(tasks) + 1):
        if not tasks[day - 1]:
            continue
        sheet.range(f'D{start_row + day}').value = code[day - 1]
        sheet.range(f'F{start_row + day}').value = starttime[day - 1]
        sheet.range(f'G{start_row + day}').value = endtime[day - 1]
        sheet.range(f'K{start_row + day}').value = tasks[day - 1]


def get_vacation_in_month(vacations, month, year):
    vacations_in_month = []
    for vacation in vacations:
        vacation_date = datetime.strptime(vacation['date'], '%Y/%m/%d')
        if vacation_date.month == month and vacation_date.year == year:
            vacations_in_month.append({'date': vacation_date, 'type': vacation['type']})
    return vacations_in_month


def write_vacations_day(sheet, vacations_in_month, start_row):
    for vacation in vacations_in_month:
        sheet.range(f'D{start_row + vacation['date'].day}').value = vacation['type']


def import_data(xlsm_file_path, csv_file_path):
    app = xw.App(visible=False)
    try:
        # List member
        config = open('config.json')
        try:
            data = json.load(config)
            members = data['members']
            vacations = data['vacations']
            # Parse the date format from config
            date_format_str = data['date_format']
            # Map the date components to their corresponding format codes
            format_map = {'Y': '%Y', 'YYYY': '%Y', 'm': '%m', 'mm': '%m', 'd': '%d', 'dd': '%d'}
            
            # Handle different separators (- or /)
            if '-' in date_format_str:
                components = date_format_str.split('-')
                date_format = '-'.join(format_map[component] for component in components)
            elif '/' in date_format_str:
                components = date_format_str.split('/')
                date_format = '-'.join(format_map[component] for component in components)
            else:
                # Default format if no separator found
                date_format = '%Y-%m-%d'
        except json.decoder.JSONDecodeError as e:
            logging.error(
                f"Invalid JSON format in config file: {str(e)}. "
                f"Please visit https://jsonformatter.org/ to reformat the JSON file to the correct format. ")
            raise
        except KeyError as e:
            logging.error(f"Missing required key in config file: {str(e)}")
            raise
        finally:
            config.close()

        logtimeData = logtime_data(csv_file_path)
        members_in_logtime = list(logtimeData['User'].unique())

        # Get list member exists in file config and logtime
        contained_member = [member for member in members.values() if member in members_in_logtime]
        dir_contained_member = {
            fullname: nickname for fullname, nickname in members.items() if nickname in contained_member
        }

        if not contained_member:
            raise ValueError('Can not found any member in config.json file exists in timelog file')

        xlsm_file_name = os.path.splitext(xlsm_file_path)
        updated_xlsm_file_path = f"{xlsm_file_name[0]}_update{xlsm_file_name[1]}"

        copy_xlsm_file(app, xlsm_file_path, updated_xlsm_file_path)
        wb = app.books.open(updated_xlsm_file_path)

        sheet = wb.sheets['設定']
        year = int(sheet.range('C5').value)
        month = int(sheet.range('C7').value)

        members_data_in_month = init_members_data_in_month(dir_contained_member, month, year, wb)
        members_data_in_month = generate_data_to_each_member(logtimeData, members_data_in_month, date_format)

        # Access the sheet where you want to import data
        for fullname, nickname in members.items():
            if nickname not in contained_member:
                continue
            if fullname not in [sheet.name for sheet in wb.sheets]:
                raise KeyError(f"Sheet '{fullname}' does not exist in the workbook. Skipping.")

            sheet = wb.sheets[fullname]

            # Prepare data for writing in batches
            code = []
            starttime = []
            endtime = []
            task = []

            for day in members_data_in_month[nickname]:
                code.append([day[1]])
                starttime.append([day[2]])
                endtime.append([day[3]])
                task.append([", ".join(day[4])] if day[4] != [] else [])

            task = [list(np.unique(np.array(value))) for value in task]

            # Write data in batches
            write_data(sheet, code, starttime, endtime, task, start_row=9)

            # Edit vacation day
            vacations_in_month = get_vacation_in_month(vacations, month, year)
            write_vacations_day(sheet, vacations_in_month, start_row=9)

        # Save and close the Excel file
        wb.save()
        wb.close()

    except KeyError as e:
        logging.error(str(e))
        raise
    except FileNotFoundError as e:
        logging.error(f"File not found: {str(e)}")
        raise
    except xw.XlwingsError as e:
        logging.error(f"An error occurred while interacting with Excel: {str(e)}")
        raise
    except PermissionError as e:
        logging.error(f"Permission denied: {str(e)}")
        raise
    except Exception as e:
        logging.error(f"An error occurred during import: {str(e)}")
        raise
    finally:
        app.quit()


# Configure logging
logging.basicConfig(filename='app.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')
