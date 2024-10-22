#!/usr/bin/env python

"""
    Ramadan dates
"""

__author__      = "Arshad H. Siddiqui"
__copyright__   = "Free to all"

from hijridate import Hijri, Gregorian
from datetime import datetime, timedelta

# Convert a Hijri date to Gregorian
#g = Hijri(1403, 2, 17).to_gregorian()

# Convert a Gregorian date to Hijri
#h = Gregorian(1982, 12, 2).to_hijri()

def get_ramadan_dates(year):

    current_date = datetime(year, 1, 1)
    while current_date.year == year:
        m = Gregorian(current_date.year, current_date.month, current_date.day).to_hijri()
        current_date += timedelta(days=1)
        if (m.month == 9 and m.day == 1):
            ramadan_start_date = current_date
        if (m.month == 10 and m.day == 1):
            ramadan_end_date = current_date-timedelta(days=-1)
    return ramadan_start_date, ramadan_end_date
