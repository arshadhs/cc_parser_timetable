#!/usr/bin/env python

"""
    Salah time logic
"""

__author__      = "Mohammad Azim Khan, Arshad H. Siddiqui"
__copyright__   = "Free to all"

import datetime
import argparse
import math

from dateutil import relativedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from ramadan_dates import get_ramadan_dates
from configReader import get_config


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
        return self.set_jamat_time(hour, min)

    def set_jamat_time(self, hour, min):

        # Jum'ah (Friday and Dhuhr) - jamat_time
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

        # "Maghrib" - jamat_time
        # Before 6 pm, then no congregation or booking
        # Ramadan, then no congregation or booking
        if self.name == "Maghrib":

            if (self.date >= self.ramadan_start and self.date <= self.ramadan_end):
                return ""
            if(int(hour) < 18):
                return ""

        # Isha - jamat_time
        # If the start time is before 19:51, set to 20:05:00
        if self.name == "Isha":
            if (int(hour) == 19 and min <= 50) or int(hour) < 19:
                    return datetime.time(20, 5, 00)
            else:
                # Start of next quarter of the hour
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

        # Fajr - jamat_time
        if self.name == "Fajr":
            fajr_hour = int(self.sunrise.strftime('%H'))
            fajr_min = int(self.sunrise.strftime('%M'))
            
            # Fajr - Ramadan
            # Start of next quarter of the hour, except on first date of Ramadan (i.e. the Fajr before Ramadan start)
            if (self.date >= self.ramadan_start and self.date <= self.ramadan_end and self.ramadan_start != self.date):
                tm = (datetime.datetime.combine(datetime.date(1,1,1), self.start) + datetime.timedelta(minutes = 15)).time()
                # print("Ramadan: Booking time: ", self.date, self.name, tm)
                return tm

            if int(hour) == 6 and int(min) > 15:
                #if (self.date.strftime('%m') == "02"): print("Warning: Booking time: ", self.date, self.name)
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

        #print("Warning: Booking time: ", self.name)

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


# If it is Saturday, pick the max from the week so as not to have a Jamat before the start time
# Use the Jamat time (and booking time) this week from Thursday
#   Do not change if it is Ramadan dates
#   Be careful when in the week with Day Light Saving transitioning
def recalculate_jamat_time(salahTable, dstDates):
#    print(type(salahTable))
#    print("time.values(): ", type(salahTable.keys()), " : ", salahTable.values())
    

    for row, (date, time) in enumerate(salahTable['schedule'].items(), start=1): # For each day in the year
#        print ("date: ", date)      # date:  2025-12-29
#        print ("time: ", time)      # time:  OrderedDict({
                                    # 'Fajr': <salah.Salah object at 0x04639DB0>, 'Sunrise': datetime.time(8, 9), 
                                    # 'Dhuhr': <salah.Salah object at 0x0463C198>, 'Asr': <salah.Salah object at 0x0463C2A0>,
                                    # 'Maghrib': <salah.Salah object at 0x0463C3C0>, 'Isha': <salah.Salah object at 0x0463C4C8>})

        # Ramadan : do not recalculate
        #if (time['Fajr'].date >= time['Fajr'].ramadan_start and time['Fajr'].date <= time['Fajr'].ramadan_end and int(time['Fajr'].ramadan_start.strftime('%d')) != int(time['Fajr'].date.strftime('%d'))):
        if (time['Fajr'].date >= time['Fajr'].ramadan_start and time['Fajr'].date <= time['Fajr'].ramadan_end and time['Fajr'].ramadan_start != time['Fajr'].date):
            continue

        # Day is Saturday : get max Jamat time from today to Friday and reset the week with the max
        if (date.weekday() == 5):
            findMaxJamatTime(salahTable, date)

        # Day is Friday : do not recalculate
        if (date.weekday() == 4):
            continue

        #d = salahTable.values()
        todayFajr, todaySunrise, todayDhuhr, todayAsr, todayMaghrib, todayIsha = time.values()

        # thursdayDate = date + relativedelta.relativedelta(weekday=3)
        # thursdaySalah = getSalahObject(salahTable, thursdayDate)
        fridayDate = date + relativedelta.relativedelta(weekday=4)
        fridaySalah = getSalahObject(salahTable, fridayDate)

        sundayDate = date + relativedelta.relativedelta(weekday=6)
        saturdaySalah = None
        if sundayDate in dstDates:
            saturdayDate = date + relativedelta.relativedelta(weekday=5)
            saturdaySalah = getSalahObject(salahTable, saturdayDate)

        #print(s)
        # if (thursdaySalah):
            # Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = thursdaySalah.values()
            # time['Fajr'].jamat = thursdaySalah['Fajr'].jamat

        if (saturdaySalah): # DST Transition
            Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = saturdaySalah.values()
            time['Fajr'].jamat = saturdaySalah['Fajr'].jamat
            time['Fajr'].booking_start = saturdaySalah['Fajr'].booking_start
            time['Fajr'].booking_end = saturdaySalah['Fajr'].booking_end
            time['Isha'].jamat = saturdaySalah['Isha'].jamat
            time['Isha'].booking_start = saturdaySalah['Isha'].booking_start
            time['Isha'].booking_end = saturdaySalah['Isha'].booking_end
        elif (fridaySalah): # Get the Jamat and booking time for coming Friday and use it
            Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = fridaySalah.values()
            time['Fajr'].jamat = fridaySalah['Fajr'].jamat
            time['Fajr'].booking_start = fridaySalah['Fajr'].booking_start
            time['Fajr'].booking_end = fridaySalah['Fajr'].booking_end
            time['Isha'].jamat = fridaySalah['Isha'].jamat
            time['Isha'].booking_start = fridaySalah['Isha'].booking_start
            time['Isha'].booking_end = fridaySalah['Isha'].booking_end

#        print(fridayDate)
#        print("time.values(): ", time.keys(), " : ", time.values())
#        print(d['datetime.date(date)'])
    return salahTable


# Every Saturday, check for Jamat times till Friday,
# find the largest and reset the whole week (Sat to Fri) Jamat time (and booking start and end)
def findMaxJamatTime(salahTable, startDate):
    saturdayDate = startDate #+ relativedelta.relativedelta(weekday=5)
    saturdaySalah = getSalahObject(salahTable, saturdayDate)
    maxJamatTime = saturdaySalah['Isha'].jamat
    booking_start = saturdaySalah['Isha'].booking_start
    booking_end = saturdaySalah['Isha'].booking_end

    for i in [5, 0, 1, 2, 3, 4]:
        date = startDate + relativedelta.relativedelta(weekday=i)
        dateSalah = getSalahObject(salahTable, date)
        if (dateSalah):
            if maxJamatTime < dateSalah['Isha'].jamat:
                maxJamatTime = dateSalah['Isha'].jamat
                booking_start = dateSalah['Isha'].booking_start
                booking_end = dateSalah['Isha'].booking_end
            else:
                dateSalah['Isha'].jamat = maxJamatTime
                dateSalah['Isha'].booking_start = booking_start
                dateSalah['Isha'].booking_end = booking_end


def getSalahObject(salahTable, jumpDate):
    for row, (date, time) in enumerate(salahTable['schedule'].items(), start=1):
        if (date == jumpDate):
            return time
