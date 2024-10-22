#!/usr/bin/env python
r"""
    Implements Logic to build up a Salah Planner using data from moonsighting (url or xlsx)
    
    Usage:
    python main.py --year 2025 --file docs\salah2025.xlsx --usage web
    C:\GitHub\cc\salah>python main.py --year 2025 --file docs\salah2025.xlsx --usage booking hide
"""

__author__      = "Mohammad Azim Khan, Arshad H. Siddiqui"
__copyright__   = "Free to all"

import datetime
import argparse
import math

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.styles.borders import Border, Side

import xlsx_writer
import csv_writer
import utils
from moon_sighting import get_prayer_table, get_prayer_table_offline
from salah_object import Salah, recalculate_jamat_time
from xlsx_generator import SalahWorkBook, FajrSalahWorkBook
from validator import validateJamatTime

COLOUR_BLUE = "add8e6"
COLOUR_GREY = "dbdbdb"
COLOR_L_GREY = "6e6e6e"

def salah_gen(table):
    """
    Builds a Salah object for each prayer time in the timetable.
    """
    for row, (date, time) in enumerate(table['schedule'].items(), start=1): # For each day in the year
        # print("time.values(): ", time.keys(), " : ", time.values())
                                    # odict_keys(['Fajr', 'Sunrise', 'Dhuhr', 'Asr', 'Maghrib', 'Isha'])
                                    # odict_values(['06:28', '08:09', '12:08', '14:10', '16:00', '17:34'])
        #print ("date: ", date)     # date:  dec 30 tue
        #print ("time: ", time)     # time:  ordereddict({'fajr': '06:28', 'sunrise': '08:09', 'dhuhr': '12:07', 'asr': '14:09', 'maghrib': '15:59', 'isha': '17:33'})


        Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = time.values()

        # Create Salah objects with specified jamat settings
        time['Fajr'] = Salah('Fajr', date, Fajr, time, has_jamat=True)
        time['Dhuhr'] = Salah('Dhuhr', date, Dhuhr, time, has_jamat=True)
        time['Asr'] = Salah('Asr', date, Asr, time, has_jamat=False)
        time['Maghrib'] = Salah('Maghrib', date, Maghrib, time, has_jamat=True)
        time['Isha'] = Salah('Isha', date, Isha, time, has_jamat=True)
#    print (table)
    return table


def salah_calculator(salahTable, dstDates):
    """
    Recalculates Jamat times for Salah objects based on DST adjustments.
    """
    salahReCalcTable = recalculate_jamat_time(salahTable, dstDates)
    return salahTable # salahReCalcTable


def workbook_gen(table):
    """
    Generates a workbook for Salah timetable data.
    """
    for row, (date, time) in enumerate(table['schedule'].items(), start=1): # For each day in the year
#        print ("date: ", date)      # date:  2025-12-29
#        print ("time: ", time)      # time:  OrderedDict({
                                    # 'Fajr': <salah.Salah object at 0x04639DB0>, 'Sunrise': datetime.time(8, 9), 
                                    # 'Dhuhr': <salah.Salah object at 0x0463C198>, 'Asr': <salah.Salah object at 0x0463C2A0>,
                                    # 'Maghrib': <salah.Salah object at 0x0463C3C0>, 'Isha': <salah.Salah object at 0x0463C4C8>})
        Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = time.values()

        time['Fajr'] = FajrSalahWorkBook(time['Fajr'], usage)
        time['Dhuhr'] = SalahWorkBook(time['Dhuhr'], usage)
        time['Asr'] = SalahWorkBook(time['Asr'], usage)
        time['Maghrib'] = SalahWorkBook(time['Maghrib'], usage)
        time['Isha'] = SalahWorkBook(time['Isha'], usage)
#    print (table)
    return table


usage = ""
hideColumns = False

# Main function to parse arguments and generate the Salah timetable
def main():
    parser = argparse.ArgumentParser(description='Generate Salah Time Table for booking and web.')
    parser.add_argument('--year', dest='year', default=datetime.datetime.now().year, help='Year of salah timetable')
    parser.add_argument('--file', dest='filename', help='Input XLS file for timetable')
    parser.add_argument('--usage', dest='usage', type=str, choices=["booking", "web"], default="web", help='Output file type, booking or web')
    parser.add_argument('hide', default=argparse.SUPPRESS, nargs='?', help='Hide certain columns (if generating booking sheet)')
    args = parser.parse_args()

    global usage
    usage = args.usage

    global hideColumns

    if hasattr(args, 'hide') and usage == "booking":
        hideColumns = True

    # ToDo: module not returning expected results
    # ramadan_start, ramadan_end = get_ramadan_dates(int(args.year))
    # print(f"Ramadan in {args.year} starts on {ramadan_start} and ends on {ramadan_end}")

    # Get Ramadan dates from config file
    ramadan_start, ramadan_end = utils.get_config("config.ini", args.year)
    print(f"Ramadan Start ({args.year}): {ramadan_start}")
    print(f"Ramadan End ({args.year}): {ramadan_end}")

    # Fetch moon sighting data from file or URL
    if args.year is not None:
        moonSightTable = get_prayer_table_offline(args.year, args.filename)
    else:
        moonSightTable = get_prayer_table(args.year)

    # Generate Salah objects
    salahBookingTable = salah_gen(moonSightTable)
    dstDates = utils.getDSTtransitionDates(int(args.year))
    salahCalculatedTable = salah_calculator(salahBookingTable, dstDates)

    # Validate calculated Jamat times
    validateJamatTime(salahCalculatedTable)

    # Generate the output file based on usage type
    if (args.usage == "booking"):
        wbTable = workbook_gen(salahCalculatedTable)
        xlsx_writer.not_in_use(args.filename)
        xlsx_writer.writer(wbTable, args.year, args.usage, hideColumns)

    if (args.usage == "web"):
        csv_writer.csvWriter(salahCalculatedTable, args.year)

if __name__ == '__main__':
    main()