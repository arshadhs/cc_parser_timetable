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

import datetime
import argparse
import math

COLOUR_BLUE = "add8e6"
COLOR_P_BLUE = "1e7ba0"
COLOR_S_BLUE = "0a2842"

COLOUR_GREY = "dbdbdb"
COLOR_L_GREY = "6e6e6e" # "ADD8E6"
COLOR_D_GREY = "474747" #"72bcd4"

class Salah(object):
    def __init__(self, name, date, start, details, has_jamat=True, fill_color=None, header_color=None):
        self.name = name
        self.date = date
        self.start = start
        self.details = details
        self.has_jamat = has_jamat
        self.fill_color = fill_color
        self.header_color = header_color
        self.sunrise = self.details['Sunrise']
        self.week_day = date.strftime('%a')
        self.ramadan_start, self.ramadan_end = get_config("config.ini", int(self.date.strftime('%Y')))
        self.location = str(self.get_location())
        self.jamat = self.get_jamat_time() # str(self.get_jamat_time())
        self.booking_start, self.booking_end = self.get_booking_time_slot()
        self.color_me = fill_color is not None
        self.is_juma = self.week_day.lower().startswith('fri')
        if self.is_juma:
            self.fill_color = PatternFill("solid", fgColor=COLOUR_BLUE)
            self.color_me = True


    # __str__ method to customize how the object is printed
    def __str__(self):
        return f"Salah(name={self.name}, date={self.date}, jamat={self.jamat})"

    def ceil_dt(self, dt, delta):
        # Calculate the number of minutes since midnight
        minutes = dt.hour * 60 + dt.minute
        
        # Always round up to the nearest delta minutes
        rounded_minutes = math.floor(minutes / delta) * delta
        
        # Calculate the new hour and minute
        new_hour = rounded_minutes / 60
        new_minute = rounded_minutes % 60
        
        # If rounding moves the time into the next day, handle the overflow
        if new_hour >= 24:
            dt = dt + timedelta(days=1)
            new_hour = 0
        
        # Return the original date with the rounded up time
        return dt.replace(hour=int(new_hour), minute=new_minute, second=0, microsecond=0)

    def get_jamat_time(self):
        hour = int(self.start.strftime('%H'))
        min = int(self.start.strftime('%M'))

        if self.name == "Dhuhr":
            if self.week_day == "Fri":
                self.location == "HUB"
                if int(hour) == 13 and min > 8:
                    return datetime.time(13, 15, 00)
                elif int(hour) == 13:
                    return datetime.time(13, 10, 00)
                else:
                    return datetime.time(13, 5, 00)
            else:
                return ""

        # If "Maghrib" is before 6 pm, then no congregation or booking
        if self.name == "Maghrib":

            if (self.date >= self.ramadan_start and self.date <= self.ramadan_end):
                return ""
            if(int(hour) < 18):
                return ""

        if self.name == "Isha":
            if (int(hour) == 19 and min <= 50) or int(hour) < 19:
                    return datetime.time(20, 00, 00)
            else:
                if min > 45:
                    new_hour = int(hour) + 1
                    if len(str(new_hour)) == 1:
                        new_hour = '0'+str(new_hour)
                    return datetime.time(new_hour, 00, 00)
                elif min > 30 and min <= 45:
                    return datetime.time(hour, 45, 00)
                elif min > 15 and min <= 30:
                    return datetime.time(hour, 30, 00)
                elif min > 0 and min <= 15:
                    return datetime.time(hour, 15, 00)
                else:
                    return(self.start)

        if self.name == "Fajr":
            fajr_hour = int(self.sunrise.strftime('%H'))
            fajr_min = int(self.sunrise.strftime('%M'))
            
            #print(f' {self.date.strftime('%m')}')
            if (self.date >= self.ramadan_start and self.date <= self.ramadan_end and int(self.ramadan_start.strftime('%d')) != int(self.date.strftime('%d'))):
                tm = (datetime.datetime.combine(datetime.date(1,1,1), self.start) + datetime.timedelta(minutes = 15)).time()
                return tm

            if int(hour) == 6 and int(min) > 15:
                return datetime.time(6, 45, 00)
            elif int(hour) == 6:
                return datetime.time(6, 30, 00)
            elif (int(hour) == 5 and (int(min) >= 45)):
                return datetime.time(6, 30, 00)
            elif int(hour) == 2:
                return datetime.time(4, 00, 00)
            elif fajr_min > 45:
                return datetime.time(fajr_hour, 00, 00)
            elif (int(fajr_min) > 30 and fajr_min <= 45):
                return datetime.time(fajr_hour - 1, 45, 00)
            elif (int(fajr_min) > 15 and fajr_min <= 30):
                return datetime.time(fajr_hour - 1, 30, 00)
            elif (int(fajr_min) >= 00 and fajr_min <= 15):
                return datetime.time(fajr_hour - 1, 15, 00)

        if min > 55:
            new_hour = int(hour) + 1
            if len(str(new_hour)) == 1:
                new_hour = '0'+str(new_hour)
            return datetime.time(new_hour, 15, 00)
        elif min > 50 and min <= 55:
            return datetime.time(hour, 55, 00)
        elif min > 45 and min <= 50:
            return datetime.time(hour, 50, 00)
        elif min > 40 and min <= 45:
            return datetime.time(hour, 45, 00)
        elif min > 35 and min <= 40:
            return datetime.time(hour, 40, 00)
        elif min > 30 and min <= 35:
            return datetime.time(hour, 35, 00)
        elif min > 25 and min <= 30:
            return datetime.time(hour, 30, 00)
        elif min > 20 and min <= 25:
            return datetime.time(hour, 25, 00)
        elif min > 15 and min <= 20:
            return datetime.time(hour, 20, 00)
        elif min > 10 and min <= 15:
            return datetime.time(hour, 15, 00)
        elif min > 5 and min <= 10:
            return datetime.time(hour, 10, 00)
        elif min > 0 and min <= 5:
            return datetime.time(hour, 5, 00)
        else:
            return(self.start)

    def get_booking_time_slot(self):

        if self.jamat == "":
            return None, None

        if self.name == "Dhuhr" and self.week_day == "Fri":
            return datetime.time(13, 00, 00), datetime.time(14, 00, 00)

        start_time = self.jamat

        if (self.date >= self.ramadan_start and self.date <= self.ramadan_end) and self.name == "Isha":
            booking_duration = 120
        else:
            booking_duration = 30

        end_time = (datetime.datetime.combine(datetime.date(1,1,1), self.jamat) + datetime.timedelta(minutes = booking_duration)).time()

        return start_time, end_time


    def get_location(self):
        if self.name == "Fajr":
            if (int(self.date.strftime('%m')) > 3 and int(self.date.strftime('%m')) < 10) and (int(self.date.strftime('%w')) == 0 or int(self.date.strftime('%w')) == 6):
                return("BS-H")
            else:
                return("NCP")
        elif self.name == "Dhuhr" and self.week_day == "Fri":
            return("HUB")
        elif self.name == "Isha" and (self.date >= self.ramadan_start and self.date <= self.ramadan_end):
            return("HUB")

        return ""

    def _style_header(self, ws, row, col):
        if self.header_color:
            ws.cell(row, col).fill = self.header_color
        ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
        ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")

    def add_xl_top_header(self, ws, row, col, do_not_merge=False):
        start_col = col
        ws.cell(row, col).value = self.name                     # xlsx: Fajr, Dhuhr, Asr, Maghrib, Isha
        self._style_header(ws, row, col)
        col += 1
        if self.has_jamat:
            col += 3 # jamat, booking and location columns
        if not do_not_merge:
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=col - 1)
        return col

    def add_xl_header(self, ws, row, col):
        ws.cell(row, col).value = 'Start'
        self._style_header(ws, row, col)
        col += 1
        if self.has_jamat:
            ws.cell(row, col).value = 'Booking'
            self._style_header(ws, row, col)
            col += 1
            ws.cell(row, col).value = 'Jamat'
            self._style_header(ws, row, col)
            col += 1
            ws.cell(row, col).value = 'Location'
            self._style_header(ws, row, col)
            col += 1
        return col

    # Fill the we.cell(row,col). value
    def add_xl_columns(self, ws, row, col):
        ws.cell(row, col).value = displayTime(self.start)  # Start Time
        if self.color_me:
            ws.cell(row, col).fill = self.fill_color
        if self.is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        if self.has_jamat:
            ws.cell(row, col).value = displayTime(self.booking_start) + "-" + displayTime(self.booking_end) if self.booking_start else ""
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            if self.is_juma:
                ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = displayTime(self.jamat) if self.jamat else "" # Jamat
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            if self.is_juma:
                ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = self.location
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            if self.is_juma:
                ws.cell(row, col).font = Font(bold=True)
            col += 1
        return col


# To add Sunrise column
class FajrSalah(Salah):
    def __init__(self, date, start, details, has_jamat=True, fill_color=None, header_color=None):
        super().__init__('Fajr', date, start, details, has_jamat=has_jamat, fill_color=fill_color, header_color=header_color)
        self.sunrise = details['Sunrise']

    def add_xl_top_header(self, ws, row, col):
        start_col = col
        col = super().add_xl_top_header(ws, row, col, do_not_merge=True)
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=col)
        col += 1
        return col

    def add_xl_header(self, ws, row, col):
        col = super().add_xl_header(ws, row, col)
        ws.cell(row, col).value = 'Sunrise'
        self._style_header(ws, row, col)
        col += 1
        return col

    # Fill the we.cell(row,col). value with sunrise time
    def add_xl_columns(self, ws, row, col):
        col = super().add_xl_columns(ws, row, col)
        ws.cell(row, col).value = displayTime(self.sunrise)
        if self.color_me:
            ws.cell(row, col).fill = self.fill_color
        if self.is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        return col


def salah_org(table):
    fill_grn = PatternFill("solid", fgColor=COLOUR_GREY)
    header_color = PatternFill("solid", fgColor=COLOR_L_GREY)

    for row, (date, time) in enumerate(table['schedule'].items(), start=1):
        fill_color = fill_grn if row % 2 != 0 else None

        # print("time.values(): ", time.keys(), " : ", time.values())
        # odict_keys(['Fajr', 'Sunrise', 'Dhuhr', 'Asr', 'Maghrib', 'Isha'])
        # odict_values(['06:28', '08:09', '12:08', '14:10', '16:00', '17:34'])

        Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = time.values()

        # print ("date: ", date)    # date:  Dec 30 Tue
        # print ("time: ", time)    # time:  OrderedDict({'Fajr': '06:28', 'Sunrise': '08:09', 'Dhuhr': '12:07', 'Asr': '14:09', 'Maghrib': '15:59', 'Isha': '17:33'})

        # Build a Salah Object for time['x']
        # it swaps out time['x'] from 'time' to 'Salah' object
        time['Fajr'] = FajrSalah(date, Fajr, time, fill_color=fill_color, header_color=header_color)
        time['Dhuhr'] = Salah('Dhuhr', date, Dhuhr, time, has_jamat=True, fill_color=fill_color, header_color=header_color)
        time['Asr'] = Salah('Asr', date, Asr, time, has_jamat=False, fill_color=fill_color, header_color=header_color)
        time['Maghrib'] = Salah('Maghrib', date, Maghrib, time, has_jamat=True, fill_color=fill_color, header_color=header_color)
        time['Isha'] = Salah('Isha', date, Isha, time, has_jamat=True, fill_color=fill_color, header_color=header_color)

    #print (table)
    return table


def generate_xl(table, year):
    table = salah_org(table)
    wb = Workbook()
    wb['Sheet'].title = 'CC booking'
    ws = wb['CC booking']
    fill_color = PatternFill("solid", fgColor=COLOUR_GREY)
    header_color = PatternFill("solid", fgColor=COLOR_L_GREY)
    row = 1

    # Add each day as rows
    for serial_no, (date, day) in enumerate(table['schedule'].items(), start=1):

        week_day = date.strftime('%a')
        is_juma = week_day == 'Fri'
        if is_juma:
            fill_color = PatternFill("solid", fgColor=COLOUR_BLUE)
        else:
            fill_color = PatternFill("solid", fgColor=COLOUR_GREY)
        color_me = fill_color if is_juma or serial_no % 2 != 0 else None
        
        # Add Salah header

        # First day of the year and web
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
                if isinstance(salah, Salah):
                    col = salah.add_xl_header(ws, row, col)
            row += 1

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
                if isinstance(salah, Salah):
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
                if isinstance(salah, Salah):
                    col = salah.add_xl_header(ws, row, col)
            row += 1

        # Fill up the actual data
        col = 1
        ws.cell(row, col).value = serial_no                     # Serial Number
        serial_no += 1
        if color_me:
            ws.cell(row, col).fill = fill_color
        if is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        ws.cell(row, col).value = date.strftime('%b-%d') if usage == "booking" else date.strftime('%d/%m/%Y')       # Date
        if color_me:
            ws.cell(row, col).fill = fill_color
        if is_juma:
            ws.cell(row, col).font = Font(bold=True)

        if usage == "booking":
            col += 1
            ws.cell(row, col).value = date.strftime('%a')           # Day

        if color_me:
            ws.cell(row, col).fill = fill_color
        if is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        for salah in day.values():
            if isinstance(salah, Salah):
                col = salah.add_xl_columns(ws, row, col)
        row += 1

    outFile = 'I:\\My Drive\\temp\\Cambourne_salah_timetable_'+year+'.xlsx'
    wb.save(outFile)
    print("\nWritten to", outFile)
    # if not_in_use(outFile):
        # wb.save(outFile)
        # print("\nWritten to", outFile)
    # else:
        # print("\nError[13]: Permission denied", outFile)


# def not_in_use(filename):
        # try:
            # os.rename(filename,filename)
            # return True
        # except:
            # return False

# date format
def displayTime(time):
    return time.strftime("%H:%M") if usage == "booking" else time.strftime("%H:%M:%S")

usage = ""

def main():
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('--year', dest='year',
                    default=datetime.datetime.now().year, help='Year of salah timetable')
    parser.add_argument('--file', dest='filename',
                    help='XLS file for timetable')
    parser.add_argument('--usage', dest='usage', type=str, choices=["booking", "web"],
                    default="web", help='booking or web')
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
