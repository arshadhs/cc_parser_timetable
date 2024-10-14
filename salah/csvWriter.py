#!/usr/bin/env python

"""
    xlsx writer
"""

__author__      = "Arshad H. Siddiqui"
__copyright__   = "Free to all"

import datetime
import math
import csv

import salahUtils

from moonsighting import get_prayer_table, get_prayer_table_offline
from salah import Salah, recalculate_jamat_time
from salahWorkBookGen import SalahWorkBook, FajrSalahWorkBook
from validator import validateJamatTime

def csvWriter(wbTable, year):

    #print(wbTable)

    outFile = 'Cambourne_web_'+year+'.csv'

    header = ['id', 'date', 'fajr', 'sunrise', 'fajr_con', 'fajr_loc', 'dhuhr', 'dhuhr_con', 'dhuhr_loc', 'asr', 'asr_con', 'asr_loc', 'maghrib', 'maghrib_con', 'maghrib_loc', 'isha', 'isha_con', 'isha_loc', 'arabic_date_text']
    # Write dictionary to CSV file
    with open(outFile, 'w', newline='') as csvfile:
        fieldnames = ['Name', 'Age', 'City']
        writer = csv.DictWriter(csvfile, fieldnames=header)
        writer.writeheader()

        data = {}

        # # Add each day as rows
        for serial_no, (date, salahData) in enumerate(wbTable['schedule'].items(), start=1):
            #print(salahData)
            # week_day = date.strftime('%a')
            # is_juma = week_day == 'Fri'


            # data - fill up the actual data
            #col = 1
            data['id'] = serial_no                     # id (serial number)
            serial_no += 1

            #col += 1                                                # date
            data['date'] = date.strftime('%Y-%m-%d')
#            print(data['date'])

            for salahK, salahV in salahData.items():
#                print (salahK, salahV)
                if (salahK != "Sunrise"):
                    data[salahK.lower()] = salahV.start
                    data[salahK.lower() + '_con'] = salahV.jamat
                    data[salahK.lower() + '_loc'] = salahV.location
                elif (salahK == "Sunrise"):
                    data[salahK.lower()] = salahV

            writer.writerow(data)
            
    print("\nWritten to", outFile)


    # Values (we.cell(row,col)) - Start, Booking, Jamat, Location (Fill the value and style / colour etc.)
    def add_xl_values(ws, row, col):
        # Start
        ws.cell(row, col).value = self.displayTime(self.salah.start)  # Start Time
        ws.cell(row, col).fill = fill_color
        ws.cell(row, col).font = Font(bold=True) if self.salah.is_juma else Font(bold=False)
        col += 1

        # If there's congregation - Booking, Jamat, Location
#        if self.salah.has_jamat:
        # Booking
        if self.usage == "booking":
            ws.cell(row, col).value = self.displayTime(self.salah.booking_start) + " - " + self.displayTime(self.salah.booking_end) if self.salah.booking_start else ""
            ws.cell(row, col).fill = fill_color
            ws.cell(row, col).font = Font(bold=True) if self.salah.is_juma else Font(bold=False)
            col += 1

        # Jamat
        ws.cell(row, col).value = self.displayTime(self.salah.jamat) if self.salah.jamat else "" # Jamat
        ws.cell(row, col).fill = fill_color
        ws.cell(row, col).font = Font(bold=True) if self.salah.is_juma else Font(bold=False)
        col += 1

        # Location
        ws.cell(row, col).value = self.salah.location
        ws.cell(row, col).fill = fill_color
        ws.cell(row, col).font = Font(bold=True) if self.salah.is_juma else Font(bold=False)
        col += 1

        return col
