#!/usr/bin/env python

"""
    Implements Logic to build up a Salah Planner using data from moonsighting (url or xlsx)
"""

__author__      = "Mohammad Azim Khan, Arshad H. Siddiqui"
__copyright__   = "Free to all"

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from moonsighting import get_prayer_table, get_prayer_table_offline
from ramadan_dates import get_ramadan_dates
from configReader import get_config
from salah import Salah
from salahWorkBookGen import SalahWorkBook, FajrSalahWorkBook

import datetime
import argparse
import math

COLOUR_BLUE = "add8e6"
# # COLOR_P_BLUE = "1e7ba0"
# # COLOR_S_BLUE = "0a2842"

COLOUR_GREY = "dbdbdb"
COLOR_L_GREY = "6e6e6e" # "ADD8E6"
# # COLOR_D_GREY = "474747" #"72bcd4"


def salah_org(table):
    for row, (date, time) in enumerate(table['schedule'].items(), start=1):
        # print("time.values(): ", time.keys(), " : ", time.values())
        # odict_keys(['Fajr', 'Sunrise', 'Dhuhr', 'Asr', 'Maghrib', 'Isha'])
        # odict_values(['06:28', '08:09', '12:08', '14:10', '16:00', '17:34'])

        Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = time.values()

        #print ("date: ", date)    # date:  dec 30 tue
        #print ("time: ", time)    # time:  ordereddict({'fajr': '06:28', 'sunrise': '08:09', 'dhuhr': '12:07', 'asr': '14:09', 'maghrib': '15:59', 'isha': '17:33'})

        # Build a Salah Object for time['x']
        # it swaps out time['x'] from 'time' to 'Salah' object
        time['Fajr'] = Salah('Fajr', date, Fajr, time, has_jamat=True)
        time['Dhuhr'] = Salah('Dhuhr', date, Dhuhr, time, has_jamat=True)
        time['Asr'] = Salah('Asr', date, Asr, time, has_jamat=False)
        time['Maghrib'] = Salah('Maghrib', date, Maghrib, time, has_jamat=True)
        time['Isha'] = Salah('Isha', date, Isha, time, has_jamat=True)

        # Build WorkBook Cells for each prayer time
        time['Fajr'] = FajrSalahWorkBook(time['Fajr'], usage)
        time['Dhuhr'] = SalahWorkBook(time['Dhuhr'], usage)
        time['Asr'] = SalahWorkBook(time['Asr'], usage)
        time['Maghrib'] = SalahWorkBook(time['Maghrib'], usage)
        time['Isha'] = SalahWorkBook(time['Isha'], usage)

#    print (table)
    return table


def generate_xl(table, year):
    table = salah_org(table)
    wb = Workbook()
    wb['Sheet'].title = 'CC booking'
    ws = wb['CC booking']

    header_color = PatternFill("solid", fgColor=COLOR_L_GREY)
    row = 1

    # Add each day as rows
    for serial_no, (date, day) in enumerate(table['schedule'].items(), start=1):

        week_day = date.strftime('%a')
        is_juma = week_day == 'Fri'

        # Add Salah header

        # Header - web (2 rows every month)
        if (date.day == 1) and (date.month == 1) and usage == "web":
            # id, date, day
            col = 1
            ws.cell(row, col).value = 'id'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1
            ws.cell(row, col).value = 'date'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1

            # Start, Booking, Jamat, Location (Sunrise)
            for salah in next(iter(table['schedule'].values())).values():
                if isinstance(salah, SalahWorkBook):
                    col = salah.add_xl_header(ws, row, col)
            row += 1

        # Header - booking (1 row on top of the sheet)
        elif int(date.strftime('%d') == "01") and usage == "booking":
            # First day of the month

            # Add top header
            col = 4
            ws.cell(row, 1).value = '{} {}'.format(date.strftime('%B'), date.strftime('%y'))    # Month and Year
            ws.cell(row, 1).font = Font(bold=True, color='FFFFFF')
            ws.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row, 1).fill = header_color
            ws.cell(row, 2).fill = header_color
            ws.cell(row, 3).fill = header_color
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col - 1)

            # Fajr, Dhuhr, Asr, Maghrib, Isha
            for salah in next(iter(table['schedule'].values())).values():
                if isinstance(salah, SalahWorkBook):
                    col = salah.add_xl_top_header(ws, row, col)

            # Add Salah header

            # id, date, day
            row += 1
            col = 1
            ws.cell(row, col).value = 'id'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1
            ws.cell(row, col).value = 'date'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1

            ws.cell(row, col).value = 'day'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1

            # Start, Booking, Jamat, Location (Sunrise)
            for salah in next(iter(table['schedule'].values())).values():
                if isinstance(salah, SalahWorkBook):
                    col = salah.add_xl_header(ws, row, col)
            row += 1

        # Data - fill up the actual data
        fill_color = PatternFill("solid", fgColor=COLOUR_BLUE) if is_juma else PatternFill("solid", fgColor=COLOUR_GREY) if row % 2 != 0 else PatternFill(fgColor="ffffff")
        col = 1
        ws.cell(row, col).value = serial_no                     # id (Serial Number)
        serial_no += 1
        ws.cell(row, col).fill = fill_color
        ws.cell(row, col).font = Font(bold=True) if is_juma else Font(bold=False)

        col += 1                                                # Date
        ws.cell(row, col).value = date.strftime('%b-%d') if usage == "booking" else date.strftime('%d/%m/%Y')
        ws.cell(row, col).fill = fill_color
        ws.cell(row, col).font = Font(bold=True) if is_juma else Font(bold=False)

        if usage == "booking":                                  # Day
            col += 1
            ws.cell(row, col).value = date.strftime('%a')
            ws.cell(row, col).fill = fill_color
            ws.cell(row, col).font = Font(bold=True) if is_juma else Font(bold=False)

        col += 1                                                # Values - Start, Booking, Jamat, Location (Fill the value and style / colour etc.)
        for salah in day.values():
            if isinstance(salah, SalahWorkBook):
                col = salah.add_xl_columns(ws, row, col)
        row += 1

    setCellWidth(ws)

    outFile = 'Cambourne_salah_timetable_'+year+'.xlsx'
    wb.save(outFile)
    print("\nWritten to", outFile)
    # if not_in_use(outFile):
        # wb.save(outFile)
        # print("\nWritten to", outFile)
    # else:
        # print("\nError[13]: Permission denied", outFile)

def not_in_use(filename):
        try:
            os.rename(filename,filename)
            return True
        except:
            return False

def setCellWidth(ws):
    dims = {}
    for row in ws.rows:
        numOfCol = 0
        for cell in row:
            numOfCol = 5
            from openpyxl.cell import MergedCell
            if cell.value and not isinstance(cell, MergedCell):
                cell.alignment = Alignment(horizontal='center')
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value + 5



usage = ""

def main():
    parser = argparse.ArgumentParser(description='Generate Salah Time Table for booking and web.')
    parser.add_argument('--year', dest='year', default=datetime.datetime.now().year, help='Year of salah timetable')
    parser.add_argument('--file', dest='filename', help='XLS file for timetable')
    parser.add_argument('--usage', dest='usage', type=str, choices=["booking", "web"], default="web", help='output file type, booking or web')
    args = parser.parse_args()

    global usage
    usage = args.usage

    # ToDo: module not returning expected results
    # ramadan_start, ramadan_end = get_ramadan_dates(int(args.year))
    # print(f"Ramadan in {args.year} starts on {ramadan_start} and ends on {ramadan_end}")

    # using config file
    ramadan_start, ramadan_end = get_config("config.ini", args.year)
    print(f"Ramadan in {args.year} starts on {ramadan_start} and ends on {ramadan_end}")

    # If the filename is supplied fetch the data from file, else from URL
    if args.filename is not None:
        table = get_prayer_table_offline(args.year, args.filename)
    else:
        table = get_prayer_table(args.year)

    # If the filename is supplied fetch the data from file, else from URL
    if args.year is not None:
        table = get_prayer_table_offline(args.year, args.filename)
    else:
        table = get_prayer_table(args.year)

    generate_xl(table, args.year)

if __name__ == '__main__':
    main()
