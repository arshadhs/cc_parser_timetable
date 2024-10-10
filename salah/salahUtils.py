#!/usr/bin/env python

"""
    Salah Utilities
"""

__author__      = "Arshad H. Siddiqui"
__copyright__   = "Free to all"

import configparser
import math

from datetime import datetime, time, timedelta, date
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

def add_and_ceil_dt(dt, add_minutes, delta):
        # Reduce the time by the specified reduction in minutes
        add_time = increment_time_by_minutes_dt(dt, add_minutes)

        # Calculate the number of minutes since midnight
        minutes = add_time.hour * 60 + add_time.minute
        
        # Always round up to the nearest delta minutes
        rounded_minutes = math.ceil(minutes / delta) * delta
        
        # Calculate the new hour and minute
        new_hour = rounded_minutes // 60
        new_minute = rounded_minutes % 60

        # Return the original date with the rounded up time
        final_time = dt.replace(hour=int(new_hour), minute=new_minute, second=0, microsecond=0)
        #print("reduce_and_floor_dt: ", dt, final_time)
        return final_time

def reduce_and_floor_dt(dt, reduction_minutes, delta):
        # Reduce the time by the specified reduction in minutes
        reduced_time = reduce_time_by_minutes_dt(dt, reduction_minutes)

        # Calculate the number of minutes since midnight for the reduced time
        minutes = reduced_time.hour * 60 + reduced_time.minute
        
        # Round down to the nearest delta minutes
        rounded_minutes = math.floor(minutes / delta) * delta

        # Calculate the new hour and minute
        new_hour = rounded_minutes // 60
        new_minute = rounded_minutes % 60
        
        # Return the reduced date with the rounded down time
        final_time = reduced_time.replace(hour=int(new_hour), minute=int(new_minute), second=0, microsecond=0)
        #print("reduce_and_floor_dt: ", dt, final_time)
        return final_time
        
def reduce_time_by_minutes_dt(dt, reduction_minutes):
    reduced_time = (datetime.combine(date(1,1,1), dt) - timedelta(minutes = reduction_minutes)).time()
    return reduced_time

def increment_time_by_minutes_dt(dt, reduction_minutes):
    reduced_time = (datetime.combine(date(1,1,1), dt) + timedelta(minutes = reduction_minutes)).time()
    return reduced_time

def diff_in_minutes(dt1, dt2):
   # convert into datetimes
   dt1 = datetime.combine(datetime.now(), dt1)
   dt2 = datetime.combine(datetime.now(), dt2)
   # compute difference
   delta = dt1 - dt2        # <--- datetime.timedelta(seconds=38417)
   diff_in_minutes = int(delta.total_seconds()//60)
   return diff_in_minutes
