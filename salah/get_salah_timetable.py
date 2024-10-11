#!/usr/bin/env python

"""
    Implements Logic to build up a Salah Planner using data from moonsighting (url or xlsx)
    
    e.g.
    python get_salah_timetable.py --year 2025 --file docs\salah2025.xlsx --usage web hide
"""

__author__      = "Mohammad Azim Khan, Arshad H. Siddiqui"
__copyright__   = "Free to all"

import datetime
import argparse
import math

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.styles.borders import Border, Side

import salahUtils

from moonsighting import get_prayer_table, get_prayer_table_offline
from salah import Salah, recalculate_jamat_time
from salahWorkBookGen import SalahWorkBook, FajrSalahWorkBook
from validator import validateJamatTime
import xlsxWriter #import writeXlsx
import csvWriter

COLOUR_BLUE = "add8e6"
# COLOR_P_BLUE = "1e7ba0"
# COLOR_S_BLUE = "0a2842"

COLOUR_GREY = "dbdbdb"
COLOR_L_GREY = "6e6e6e"
# COLOR_D_GREY = "474747"

# Build a Salah Object for time['x']
# it swaps out time['x'] from 'time' to 'Salah' object
def salah_gen(table):
    for row, (date, time) in enumerate(table['schedule'].items(), start=1): # For each day in the year
        # print("time.values(): ", time.keys(), " : ", time.values())
                                    # odict_keys(['Fajr', 'Sunrise', 'Dhuhr', 'Asr', 'Maghrib', 'Isha'])
                                    # odict_values(['06:28', '08:09', '12:08', '14:10', '16:00', '17:34'])
        #print ("date: ", date)     # date:  dec 30 tue
        #print ("time: ", time)     # time:  ordereddict({'fajr': '06:28', 'sunrise': '08:09', 'dhuhr': '12:07', 'asr': '14:09', 'maghrib': '15:59', 'isha': '17:33'})


        Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = time.values()

        time['Fajr'] = Salah('Fajr', date, Fajr, time, has_jamat=True)
        time['Dhuhr'] = Salah('Dhuhr', date, Dhuhr, time, has_jamat=True)
        time['Asr'] = Salah('Asr', date, Asr, time, has_jamat=False)
        time['Maghrib'] = Salah('Maghrib', date, Maghrib, time, has_jamat=True)
        time['Isha'] = Salah('Isha', date, Isha, time, has_jamat=True)
#    print (table)
    return table


def salah_calculator(salahTable, dstDates):
    salahReCalcTable = recalculate_jamat_time(salahTable, dstDates)
    return salahTable # salahReCalcTable


# Build WorkBook Cells for each prayer time
def workbook_gen(table):
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

    # using config file
    ramadan_start, ramadan_end = salahUtils.get_config("config.ini", args.year)
    print(f"Ramadan in {args.year} starts on {ramadan_start} and ends on {ramadan_end}")

    # If the filename is supplied fetch the data from file, else from URL
    if args.filename is not None:
        moonSightTable = get_prayer_table_offline(args.year, args.filename)
    else:
        table = get_prayer_table(args.year)

    # If the filename is supplied fetch the data from file, else from URL
    if args.year is not None:
        moonSightTable = get_prayer_table_offline(args.year, args.filename)
    else:
        moonSightTable = get_prayer_table(args.year)

    # Add booking data to moon sighting data
    salahBookingTable = salah_gen(moonSightTable)
    dstDates = salahUtils.getDSTtransitionDates(int(args.year))
    salahCalculatedTable = salah_calculator(salahBookingTable, dstDates)

    # Validate the calculated booking data
    validateJamatTime(salahCalculatedTable)

    # Generate the xlsx
    wbTable = workbook_gen(salahCalculatedTable)
    xlsxWriter.not_in_use(args.filename)
    xlsxWriter.writer(wbTable, args.year, args.usage, hideColumns)

    if (args.usage == "web"):
        csvWriter.csvWriter(wbTable, args.year)

if __name__ == '__main__':
    main()
