#!/usr/bin/env python

"""
    Salah Utilities
"""

__author__      = "Arshad H. Siddiqui"
__copyright__   = "Free to all"

import configparser

from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo # Python 3.9+

# Parse the config file
def get_config(fileName, year):
    config = configparser.ConfigParser()
    config.read(fileName)

    ramadan_start = datetime.strptime(config.get(str(year), 'ramadan_start', fallback=config['DEFAULT']['ramadan_start']), '%Y-%m-%d').date()
    ramadan_end = datetime.strptime(config.get(str(year), 'ramadan_end', fallback=config['DEFAULT']['ramadan_end']), '%Y-%m-%d').date()

    return ramadan_start, ramadan_end


# Check if the date part of a datetime object falls on the date of a DST transition.
def is_date_of_DSTtransition(dt: datetime, zone: str) -> bool:
    _d = datetime.combine(dt.date(), time.min).replace(tzinfo=ZoneInfo(zone))
    return _d.dst() != (_d+timedelta(1)).dst()


# returns 2 DST Transition Dates (Sundays)
# 2025-03-30
# 2025-10-26
def getDSTtransitionDates(year):
    dstTransitionDates = []
    for d in range(366):
        if is_date_of_DSTtransition(datetime(year, 1, 1) + timedelta(d), "Europe/London"):
            dstTransitionDates.append((datetime(year, 1, 1) + timedelta(d)).date())
            print("DST Transition Sunday: ", (datetime(year, 1, 1) + timedelta(d)).date())
    return dstTransitionDates
