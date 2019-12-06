"""
Determine minimum, maximum, and mean temperatures for each plant during
the time period that the plant was monitored.
"""
from datetime import datetime

from openpyxl import load_workbook
from site_to_sheets import get_temp_sheet_name


def get_temp_file():
    """
    Loads file containing plant temperature data into openpyxl workbook object.
    :return: Openpyxl workbook object containing plant temperature data.
    """
    temp_file = load_workbook('/Users/u6000791/Box/Conservation/Rare Plants/Research Projects/Penstemon 2018-2019/'
                              'PENGRA_2019/DATA_PENGRA_2019/Measured variables_PENGRA.xlsx')
    return temp_file


def get_output_file():
    output_file = load_workbook('/Users/u6000791/Desktop/PENGRA_SEM.xlsx')
    return output_file


def get_next_plant(plants_ws):
    """
    Return row containing next plant info.
    :param plants_ws: Openpyxl worksheet containing monitored plant info.
    :return: Openpyxl row.
    """
    for row in plants_ws.iter_rows(min_row=3, max_row=43, max_col=9):
        yield row


def get_site(plant_row):
    """
    Get the site the plant is from
    :param plant_row: Openpyxl describing a monitored plant.
    :return: The string abbreviation of the site name where the plant is from.
    """
    # Return site info which is stored in the first column of the row.
    return plant_row[0].value


def get_start_time(plant_row, monitoring_date):
    """
    Get the monitoring start time for the given plant
    :param monitoring_date: A datetime object containing the day, month, and
    year that monitoring took place.
    :param plant_row: Openpyxl describing a monitored plant.
    :return: Datetime object representing when plant monitoring began.
    """
    # Start time data is in the 5th column
    return datetime(day=monitoring_date.day, month=monitoring_date.month,
                    year=monitoring_date.year, hour=plant_row[4].value.hour,
                    minute=plant_row[4].value.minute)


def get_monitoring_date(plant_row):
    """
    Get the monitoring date for the given plant
    :param plant_row: Openpyxl describing a monitored plant.
    :return: Datetime object representing the date of monitoring.
    """
    # Date value is in the 3rd column
    return plant_row[2].value


def get_end_time(plant_row, monitoring_date):
    """
    Get the monitoring end time for the given plant
    :param monitoring_date: A datetime object containing the day, month, and
    year that monitoring took place.
    :param plant_row: Openpyxl describing a monitored plant.
    :return: Datetime object representing when plant monitoring began.
    """
    # End time data is in the 6th column
    return datetime(day=monitoring_date.day, month=monitoring_date.month,
                    year=monitoring_date.year, hour=plant_row[5].value.hour,
                    minute=plant_row[5].value.minute)


def filter_temps_by_monitoring(temp_range, start_time, end_time):
    """
    Return a list of temperatures that were recorded only within the
    monitored period.
    :param temp_range: Cell range containing datetimes and recorded
    temperatures in degrees Centigrade.
    :param start_time: Datetime when monitoring began.
    :param end_time: Datetime when monitoring ended.
    :return: List of temperatures during monitoring period.
    """
    monitored_temps = []
    for row in temp_range:
        # Check if the datetime in the row's first column is in between
        # the start and end monitoring time
        temp_date = datetime.strptime(row[0].value, '%Y-%m-%d %H:%M:%S')
        if start_time <= temp_date <= end_time:
            monitored_temps.append(row[1].value)

    return monitored_temps


def main():
    # Get the workbook from our file with the plant numbers and monitoring times
    output_wb = get_output_file()

    # Get the file containing the temperature data
    temp_wb = get_temp_file()

    # Get the sheet with our plants from the output workbook
    plants_ws = output_wb['Temp']

    # Get a row from the plant worksheet
    for plant_row in get_next_plant(plants_ws):
        # Get the date of monitoring
        monitoring_date = get_monitoring_date(plant_row)

        # Get the monitoring start time for the plant
        start_time = get_start_time(plant_row, monitoring_date)

        # Get the monitoring end time for the plant
        end_time = get_end_time(plant_row, monitoring_date)

        # Determine which site the plant described in the row is from
        site = get_site(plant_row)

        # Get the name of the sheet containing temp data for this site
        temp_sheet_name = get_temp_sheet_name(site)

        # Get the actual worksheet of temp data now that we have the name
        temp_sheet = temp_wb[temp_sheet_name]

        # Get all of the datetimes and temperatures from the sheet
        temp_range = temp_sheet['B27:C1440']

        # Filter temp range by monitoring start and end times
        filtered_temp_range = filter_temps_by_monitoring(temp_range, start_time, end_time)


if __name__ == "__main__":
    main()
