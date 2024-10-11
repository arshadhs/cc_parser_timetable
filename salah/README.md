# Cambourne Salah

# Requirement 
1. python3
2. Packages required (pip3 install)
    openpyxl==3.0.10
    requests==2.28.1
    urllib3==1.26.12
    xmltodict==0.13.0
    python-certifi-win32
    hijridate
	python-dateutil
	tzdata

# Generate Salah timetable
python3.exe get_salah_timetable.py --year $year
example:
python3.exe get_salah_timetable.py --year 2023

Salah Rules 
1. prayer times are sourced from Moonsighting.com for Cambourne (52.2178° N, 0.0662° W) 
2. Nearest 5th of a minute is set for Jamat time
3. Fajr: 
        If sunrise is after 7:15 then Jamat time is set for 6:30
        If sunrise is between 6:45 and 7:15 then Jamat time is set for 6:00
        If sunrise is between 5:45 and 6:44 then Jamat time is set for 05:30
        If sunrise is between 5:30 and 5:44 then Jamat time is set for 05:00
        If sunrise is between 5:00 and 5:30 then Jamat time is set for 04:30
        If sunrise is between 4:30 and 5:00 then Jamat time is set for 04:00
        otherwise script follows the actaul start time and rounds off with nearest 5th of a minute

4. Isha: 
        If jamat time starts before 19:45, then Jamat time is set for 19:45 
        otherwise script follows the actaul start time and rounds off with nearest   5th of a minute
5. Booking time: 30 min timeslot time is set based on the Jamat start time. (Jamat start time + 30mins)
