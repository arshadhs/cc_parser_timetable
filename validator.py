#!/usr/bin/env python

"""
    Validate Data
"""

__author__      = "Arshad H. Siddiqui"
__copyright__   = "Free to all"

import datetime

import utils

def validateJamatTime(salahTable):

    print("""
Validating:
    Fajr start and jamat time difference (to allow for Sunah)
    Fajr jamat time gap from Sunrise
    Maghrib start and jamat time difference
    Isha jamat start and jamat time difference
    Jamat time is after Booking time (if Jamat)
    """)

    for row, (date, time) in enumerate(salahTable['schedule'].items(), start=1): # For each day in the year

        # Fajr jamat time difference from start
        diffInMin = utils.diff_in_minutes(time['Fajr'].jamat, time['Fajr'].start)

        if (diffInMin < 5):
           print("Error:", date, time['Fajr'].name,
                   time['Fajr'].start,
                   time['Fajr'].jamat,
                   time['Fajr'].sunrise,
                   diffInMin)
        elif (diffInMin < 15):
           print("Warning:", date, time['Fajr'].name,
                   time['Fajr'].start,
                   time['Fajr'].jamat,
                   time['Fajr'].sunrise,
                   diffInMin)



        # Fajr jamat time gap from Sunrise
        diffInMin = utils.diff_in_minutes(time['Fajr'].sunrise, time['Fajr'].jamat)

        if (diffInMin < 37):   #(utils.reduce_time_by_minutes_dt(time['Fajr'].sunrise, 37)) < time['Fajr'].jamat):
           print("Error (sunrise):", date, time['Fajr'].name,
                   time['Fajr'].start,
                   time['Fajr'].jamat,
                   time['Fajr'].sunrise,
                   utils.diff_in_minutes(time['Fajr'].sunrise, time['Fajr'].jamat))
        # elif (diffInMin < 45):   #(utils.reduce_time_by_minutes_dt(time['Fajr'].sunrise, 37)) < time['Fajr'].jamat):
           # print("Warning (sunrise):", date, time['Fajr'].name,
                   # time['Fajr'].start,
                   # time['Fajr'].jamat,
                   # time['Fajr'].sunrise,
                   # utils.diff_in_minutes(time['Fajr'].sunrise, time['Fajr'].jamat))

        # Maghrib jamat time is after start time
        if (time['Maghrib'].jamat):
            if (time['Maghrib'].start > time['Maghrib'].jamat):
               print("Error:", date, time['Maghrib'].name, time['Maghrib'].start, time['Maghrib'].jamat)

        # Isha jamat time is after start time
        if (time['Isha'].start > time['Isha'].jamat):
           print("Error:", date, time['Isha'].name, time['Isha'].start, time['Isha'].jamat)
        elif (time['Isha'].start == time['Isha'].jamat):
           print("Warning:", date, time['Isha'].name, time['Isha'].start, time['Isha'].jamat)

        # Jamat time is after Booking time
        if (time['Fajr'].booking_start > time['Fajr'].jamat):
            print("Error (Fajr Booking):", date, time['Fajr'].name, time['Fajr'].booking_start, time['Fajr'].jamat)

        if (time['Dhuhr'].jamat):
            if (time['Dhuhr'].booking_start > time['Dhuhr'].jamat):
                print("Error (Dhuhr Booking):", date, time['Dhuhr'].name, time['Dhuhr'].booking_start, time['Dhuhr'].jamat)

        if (time['Maghrib'].jamat):
            if (time['Maghrib'].booking_start > time['Maghrib'].jamat):
                print("Error (Maghrib Booking):", date, time['Maghrib'].name, time['Maghrib'].booking_start, time['Maghrib'].jamat)

        if (time['Isha'].booking_start > time['Isha'].jamat):
            print("Error (Isha Booking):", date, time['Isha'].name, time['Isha'].booking_start, time['Isha'].jamat)
