#!/usr/bin/env python

"""
    Implements Logic to build up a Salah Planner using data from moonsighting (url or xlsx)
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


def generate_xl(wbTable, year):
    wb = Workbook()
    wb['Sheet'].title = 'CC booking'
    ws = wb['CC booking']

    header_color = PatternFill("solid", fgColor=COLOR_L_GREY)
    row = 1

    # Add each day as rows
    for serial_no, (date, day) in enumerate(wbTable['schedule'].items(), start=1):

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
            for salah in next(iter(wbTable['schedule'].values())).values():
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
            for salah in next(iter(wbTable['schedule'].values())).values():
                if isinstance(salah, SalahWorkBook):
                    col = salah.add_xl_top_header(ws, row, col, hideColumns)

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
            for salah in next(iter(wbTable['schedule'].values())).values():
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

    # Hide columns not relevant for booking
    if (hideColumns):
        for col in ['D', 'F', 'H', 'I', 'K', 'I', 'M', 'N', 'P', 'R', 'T']:
            ws.column_dimensions[col].hidden= True

        thin_border = Border(left=Side(style='none'), 
                     right=Side(style='thick'), 
                     top=Side(style='none'), 
                     bottom=Side(style='none'))

        for col in ['C', 'G', 'L', 'Q', 'U']:
            for cell in ws[col]:
                cell.border = thin_border
            #ws.cell(column=2).border = thin_border

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
    generate_xl(wbTable, args.year)


if __name__ == '__main__':
    main()
