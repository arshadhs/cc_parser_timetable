#!/usr/bin/env python

"""
    Validate Data
"""

__author__      = "Arshad H. Siddiqui"
__copyright__   = "Free to all"

def validateJamatTime(salahTable):
    for row, (date, time) in enumerate(salahTable['schedule'].items(), start=1): # For each day in the year
        # Fajr jamat time is after start time
        if (time['Fajr'].start > time['Fajr'].jamat):
           print("Error:", date, time['Fajr'].name, time['Fajr'].start, time['Fajr'].jamat)
        if (time['Fajr'].start == time['Fajr'].jamat):
           print("Warning:", date, time['Fajr'].name, time['Fajr'].start, time['Fajr'].jamat)

        # Maghrib jamat time is after start time
        if (time['Maghrib'].jamat):
            if (time['Maghrib'].start > time['Maghrib'].jamat):
               print("Error:", date, time['Maghrib'].name, time['Maghrib'].start, time['Maghrib'].jamat)

        # Isha jamat time is after start time
        if (time['Isha'].start > time['Isha'].jamat):
           print("Error:", date, time['Isha'].name, time['Isha'].start, time['Isha'].jamat)
        elif (time['Isha'].start == time['Isha'].jamat):
           print("Warning:", date, time['Isha'].name, time['Isha'].start, time['Isha'].jamat)           
