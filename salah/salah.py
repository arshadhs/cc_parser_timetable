#!/usr/bin/env python

"""
    Salah time logic
"""

__author__      = "Mohammad Azim Khan, Arshad H. Siddiqui"
__copyright__   = "Free to all"

import datetime
import argparse

from datetime import timedelta
from dateutil import relativedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

import salahUtils

from ramadan_dates import get_ramadan_dates

class Salah(object):
    def __init__(self, name, date, start, details, has_jamat=True):
        self.name = name
        self.date = date
        self.start = start
        self.details = details
        self.has_jamat = has_jamat
        self.sunrise = self.details['Sunrise']
        self.week_day = date.strftime('%a')
        self.ramadan_start, self.ramadan_end = salahUtils.get_config("config.ini", int(self.date.strftime('%Y')))
        self.location = str(self.get_location())
        self.jamat = self.get_jamat_time() # str(self.get_jamat_time())
        self.booking_start, self.booking_end = self.get_booking_time_slot()
        self.is_juma = self.week_day.lower().startswith('fri')

    # __str__ method to customize how the object is printed
    def __str__(self):
        return f"Salah(name={self.name}, date={self.date}, jamat={self.jamat}, location={self.location})"

    def get_jamat_time(self):
        hour = int(self.start.strftime('%H'))
        min = int(self.start.strftime('%M'))
        return self.set_jamat_time(hour, min)

    def set_jamat_time(self, hour, min):

        # Fajr - jamat_time
        if self.name == "Fajr":
            return self.get_fajr_jamat_time(hour, min)

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

        # Asr
        if self.name == "Asr":
            return ""

        # "Maghrib" - jamat_time
        if self.name == "Maghrib":
            # Ramadan, then no congregation or booking
            if (self.date >= self.ramadan_start and self.date <= self.ramadan_end):
                return ""
            # Before 6 pm, then no congregation or booking
            # if(int(hour) < 18):
                # return ""
            if (int(hour) == 15):
                return datetime.time(16, 00, 00)

        # Isha - jamat_time
        #   If the start time is before 19:51, set to 20:05:00
        #   If time is after 22:30, add 5 minutes delta, else 15 minutes
        if self.name == "Isha":
            if (int(hour) == 19 and min <= 50) or int(hour) < 19:
                    return datetime.time(20, 5, 00)
            elif (int(hour) == 22 and min >= 30):
                    return (salahUtils.add_and_ceil_dt(self.start, 0, 5))
            else: # Start of next quarter of the hour
                return (salahUtils.add_and_ceil_dt(self.start, 0, 15))

        # Maghrib - jamat_time and other fall backs)
        #print("Fallback, why am I here? : Booking time: ", self.date, self.name)
        return (salahUtils.add_and_ceil_dt(self.start, 0, 15))

    def get_fajr_jamat_time(self, hour, min):
        fajr_sunrise_hour = int(self.sunrise.strftime('%H'))
        fajr_sunrise_min = int(self.sunrise.strftime('%M'))

#        print(dstDates)
        # Fajr - Ramadan
        # Start of next quarter of the hour, except on first date of Ramadan (i.e. the Fajr before Ramadan start)
        if (self.date >= self.ramadan_start and self.date <= self.ramadan_end and self.ramadan_start != self.date):
            tm = (datetime.datetime.combine(datetime.date(1,1,1), self.start) + datetime.timedelta(minutes = 15)).time()
            # print("Ramadan: Fajr time: ", self.date, self.name, tm)
            return tm

        # Check if the hour is 2, return the time 4:00
        if int(hour) == 2:
            #print("get_fajr_jamat_time A: ", self.date, self.start, self.sunrise, "2 am: 4 pm")
            return datetime.time(4, 0, 00)

        # Sunrise after 8:00	Time 6:40
        # Sunrise 7:50 - 8:00 	Time 6:35
        # Sunrise 7:40 - 7:50	Time 6:30
        # Sunrise 7:15 - 7:40	Time 6:15
        # Sunrise 7:00 - 7:30	Time 6:00

        elif (self.sunrise > datetime.datetime.strptime("08:00:00", "%H:%M:%S").time()):  # after 07:00
            tm = datetime.time(6, 40, 00)
            return tm

        elif ((self.sunrise <= datetime.datetime.strptime("08:00:00", "%H:%M:%S").time())
            and (self.sunrise > datetime.datetime.strptime("07:50:00", "%H:%M:%S").time())):
            tm = datetime.time(6, 35, 00)
            return tm

        elif ((self.sunrise <= datetime.datetime.strptime("07:50:00", "%H:%M:%S").time())
            and (self.sunrise > datetime.datetime.strptime("07:40:00", "%H:%M:%S").time())):
            tm = datetime.time(6, 30, 00)
            return tm

        elif ((self.sunrise <= datetime.datetime.strptime("07:40:00", "%H:%M:%S").time())
            and (self.sunrise > datetime.datetime.strptime("07:15:00", "%H:%M:%S").time())):
            tm = datetime.time(6, 15, 00)
            return tm

        elif ((self.sunrise <= datetime.datetime.strptime("07:15:00", "%H:%M:%S").time())
            and (self.sunrise > datetime.datetime.strptime("07:00:00", "%H:%M:%S").time())):
            tm = datetime.time(6, 00, 00)
            return tm

        else:
            tm = (salahUtils.reduce_and_floor_dt(self.sunrise, 46, 15))
            return tm

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

        end_time = salahUtils.increment_time_by_minutes_dt(start_time, booking_duration)
        #(datetime.datetime.combine(datetime.date(1,1,1), start_time) + datetime.timedelta(minutes = booking_duration)).time()

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
        if (time['Fajr'].date >= time['Fajr'].ramadan_start and time['Fajr'].date <= time['Fajr'].ramadan_end and time['Fajr'].ramadan_start != time['Fajr'].date):
            continue

        # Day is Saturday : get max Jamat time from today to Friday and reset the week with the max
        if (date.weekday() == 5):
            findMaxAndResetJamatTime(salahTable, date)

        # Day is Friday : do not recalculate
        if (date.weekday() == 4):
            continue

        #d = salahTable.values()
        todayFajr, todaySunrise, todayDhuhr, todayAsr, todayMaghrib, todayIsha = time.values()

        fridayDate = date + relativedelta.relativedelta(weekday=4)
        fridaySalah = getSalahObject(salahTable, fridayDate)

        sundayDate = date + relativedelta.relativedelta(weekday=6)
        saturdayDstSalah = None
        if sundayDate in dstDates:
            saturdayDstDate = date + relativedelta.relativedelta(weekday=5)
            saturdayDstSalah = getSalahObject(salahTable, saturdayDstDate)

        if (saturdayDstSalah): # DST Transition
            Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = saturdayDstSalah.values()
            if date in dstDates:
                resetDstJamatTime(salahTable, date)

            time['Fajr'].jamat = saturdayDstSalah['Fajr'].jamat
            time['Fajr'].booking_start = saturdayDstSalah['Fajr'].booking_start
            time['Fajr'].booking_end = saturdayDstSalah['Fajr'].booking_end

            time['Isha'].jamat = saturdayDstSalah['Isha'].jamat
            time['Isha'].booking_start = saturdayDstSalah['Isha'].booking_start
            time['Isha'].booking_end = saturdayDstSalah['Isha'].booking_end
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
def findMaxAndResetJamatTime(salahTable, startDate):
    saturdayDate = startDate #+ relativedelta.relativedelta(weekday=5)
    saturdaySalah = getSalahObject(salahTable, saturdayDate)

    # Fajr
    maxJamatTimeFajr = saturdaySalah['Fajr'].jamat
    booking_startFajr = saturdaySalah['Fajr'].booking_start
    booking_endFajr = saturdaySalah['Fajr'].booking_end

    # Isha
    maxJamatTime = saturdaySalah['Isha'].jamat
    booking_start = saturdaySalah['Isha'].booking_start
    booking_end = saturdaySalah['Isha'].booking_end

    for i in [5, 6, 0, 1, 2, 3, 4]:
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

            if maxJamatTimeFajr < dateSalah['Fajr'].jamat:
                maxJamatTimeFajr = dateSalah['Fajr'].jamat
                booking_startFajr = dateSalah['Fajr'].booking_start
                booking_endFajr = dateSalah['Fajr'].booking_end
            else:
                dateSalah['Fajr'].jamat = maxJamatTimeFajr
                dateSalah['Fajr'].booking_start = booking_startFajr
                dateSalah['Fajr'].booking_end = booking_endFajr


# Update the Fajr time from DST Sunday to Friday
def resetDstJamatTime(salahTable, dstSundayDate):
    # Fajr
    # jamatTimeFajr = sundaySalah['Fajr'].jamat
    # booking_startFajr = sundaySalah['Fajr'].booking_start
    # booking_endFajr = sundaySalah['Fajr'].booking_end

    sundayDate = dstSundayDate + relativedelta.relativedelta(weekday=5)
    sundaySalah = getSalahObject(salahTable, sundayDate)
#    print("resetDstJamatTime: ", dstSundayDate, sundayDate, sundaySalah['Fajr'].jamat)

    for i in [0, 1, 2, 3, 4]:
        date = dstSundayDate + relativedelta.relativedelta(weekday=i)
        dateSalah = getSalahObject(salahTable, date)

        if (dateSalah):
            dateSalah['Fajr'].jamat = sundaySalah['Fajr'].jamat
            dateSalah['Fajr'].booking_start = sundaySalah['Fajr'].booking_start
            dateSalah['Fajr'].booking_end = sundaySalah['Fajr'].booking_end
#            print("resetDstJamatTime: ", date, dateSalah['Fajr'].jamat)

def getSalahObject(salahTable, jumpDate):
    for row, (date, time) in enumerate(salahTable['schedule'].items(), start=1):
        if (date == jumpDate):
            return time