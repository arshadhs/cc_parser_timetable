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

    # # Find the approximate start date of Ramadan in the given Gregorian year
    # ramadan_start = Gregorian(year, 1, 1).to_hijri().month  # Convert January 1 of the year to Hijri to estimate Ramadan start
                    # #convert.Gregorian(year, 1, 1).to_hijri().month
    # print ("ramadan_start: ", ramadan_start)

    # # Hijri calendar is lunar, so Ramadan generally starts around 10 days earlier each year
    # ramadan_start_date = Gregorian(year, 1, 1).to_hijri()
                         # # convert.Gregorian(year, 1, 1).to_hijri().to_gregorian()
    # print ("ramadan_start_date: ", ramadan_start_date)
    # print("year_1: " , year)

    # # Iterate to find the exact start date of Ramadan
    # while True:
        # print("year: " , year)
        # hijri_date = Gregorian(year, 1, 1).to_hijri()
                        # # convert.Gregorian(year, 1, 1).to_hijri()
        # print("year: " , year)
        # if hijri_date.month == 9:  # Ramadan is the 9th month in the Hijri calendar
            # print("year: " , year)
            # ramadan_start_date = Hijri(year, 9, 1).to_gregorian()
                                    # # convert.Hijri(year, 9, 1).to_gregorian()
            # break
        # year += 1

    # # Ramadan lasts for 29 or 30 days
    # ramadan_end_date = ramadan_start_date + timedelta(days=29)
    
    # return ramadan_start_date, ramadan_end_date

# Example usage
# year = 2024
# start_date, end_date = get_ramadan_dates(year)
# print(f"Ramadan in {year} starts on {start_date} and ends on {end_date}")
