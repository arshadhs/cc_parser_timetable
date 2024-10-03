#!/usr/bin/env python

"""
    Implements Logic to build up a Salah Planner using data from moonsighting (url or xlsx)
"""

__author__      = "Mohammad Azim Khan, Arshad H. Siddiqui"
__copyright__   = "Free to all"

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from moonsighting import get_prayer_table, get_prayer_table_offline
from r_dates import get_ramadan_dates
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
    def __init__(self, name, date, start, details, has_jamat=True):
        self.name = name
        self.date = date
        self.start = start
        self.details = details
        self.has_jamat = has_jamat
        self.sunrise = self.details['Sunrise']
        self.week_day = date.strftime('%a')
        self.ramadan_start, self.ramadan_end = get_config("config.ini", int(self.date.strftime('%Y')))
        self.location = str(self.get_location())
        self.jamat = self.get_jamat_time() # str(self.get_jamat_time())
        self.booking_start, self.booking_end = self.get_booking_time_slot()
        self.is_juma = self.week_day.lower().startswith('fri')

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
                    return datetime.time(20, 5, 00)
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


        hour = int(self.jamat.strftime('%H'))
        min = int(self.jamat.strftime('%M'))
        
        if min >= 45:
            start_time = datetime.time(hour, 45, 00)
        elif min >= 30 and min < 45:
            start_time =  datetime.time(hour, 30, 00)
        elif min >= 15 and min < 30:
            start_time =  datetime.time(hour, 15, 00)
        elif min >= 0 and min < 15:
            start_time =  datetime.time(hour, 00, 00)

        if (self.date >= self.ramadan_start and self.date <= self.ramadan_end) and self.name == "Isha":
            booking_duration = 120
        else:
            booking_duration = 30

        end_time = (datetime.datetime.combine(datetime.date(1,1,1), start_time) + datetime.timedelta(minutes = booking_duration)).time()

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

    def displayTime(self, time):
        return time.strftime("%H:%M") #if usage == "booking" else time.strftime("%H:%M:%S")
