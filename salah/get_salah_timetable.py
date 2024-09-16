from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from moonsighting import get_prayer_table, get_prayer_table_offline
import datetime
import argparse

COLOUR_BLUE = "add8e6"
COLOR_P_BLUE = "1e7ba0"
COLOR_S_BLUE = "0a2842"

COLOUR_GREY = "dbdbdb"
COLOR_L_GREY = "6e6e6e" # "ADD8E6"      # HEADER
COLOR_D_GREY = "474747" #"72bcd4"

class Salah(object):
    def __init__(self, name, date, start, details, has_jamat=True, fill_color=None, header_color=None):
        self.name = name
        self.date = date
        self.start = start
        self.details = details
        self.has_jamat = has_jamat
        self.fill_color = fill_color
        self.header_color = header_color
        self.sunrise = self.details['Sunrise']
        self.jamat = str(self.get_jamat_time()) # nearest 5th of a minute
        self.booking_start = self.get_booking_time_slot()
        self.location = ''
        self.week_day = date.split(' ')[2].strip()
        self.color_me = fill_color is not None
        self.is_juma = self.week_day.lower().startswith('fri')
        if self.is_juma:
            self.fill_color = PatternFill("solid", fgColor=COLOUR_BLUE)
            self.color_me = True
    
    def get_jamat_time(self):
        print ("self: ", self.start)
        hour = self.start[:2].strip(':')
        min = int(self.start[3:6].strip(':'))

        if self.name == "Isha":
            if (int(hour) == 19 and min <= 50) or int(hour) < 19:
                    return('20:00')
            else:
                print("min: ", min)
                if min > 45:
                    new_hour = int(hour) + 1
                    if len(str(new_hour)) == 1:
                        new_hour = '0'+str(new_hour)
                    return(str(new_hour)+':00')
                elif min > 30 and min <= 45:
                    return(str(hour)+':45')
                elif min > 15 and min <= 30:
                    return(str(hour)+':30')
                elif min > 0 and min <= 15:
                    return(str(hour)+':15')
                else:
                    return(self.start[:5])

        if self.name == "Fajr":
            fajr_hour = self.sunrise[:2].strip(':')
            fajr_min = int(self.sunrise[-2:].strip(':'))
            if (int(fajr_hour) == 7 and fajr_min >= 15) or int(fajr_hour) >= 8:
                return('06:30')
            elif ((int(fajr_hour) == 7 and fajr_min < 15) or (int(fajr_hour) == 6 and fajr_min >= 45)):
                return('06:00')
            elif ((int(fajr_hour) == 6 and fajr_min < 45) or (int(fajr_hour) == 5 and fajr_min >= 45)):
                return('05:30')
            elif (int(fajr_hour) == 5 and fajr_min >= 30):
                return('05:00')
            elif (int(fajr_hour) == 5 and fajr_min < 30):
                return('04:30')
            elif (int(fajr_hour) == 4 and fajr_min >=30):
                return('04:00')

        if min > 55:
            new_hour = int(hour) + 1
            if len(str(new_hour)) == 1:
                new_hour = '0'+str(new_hour)
            return(str(new_hour)+':00')
        elif min > 50 and min <= 55:
            return(str(hour)+':55')
        elif min > 45 and min <= 50:
            return(str(hour)+':50')
        elif min > 40 and min <= 45:
            return(str(hour)+':45')
        elif min > 35 and min <= 40:
            return(str(hour)+':40')
        elif min > 30 and min <= 35:
            return(str(hour)+':35')
        elif min > 25 and min <= 30:
            return(str(hour)+':30')
        elif min > 20 and min <= 25:
            return(str(hour)+':25')
        elif min > 15 and min <= 20:
            return(str(hour)+':20')
        elif min > 10 and min <= 15:
            return(str(hour)+':15')
        elif min > 5 and min <= 10:
            return(str(hour)+':10')
        elif min > 0 and min <= 5:
            return(str(hour)+':05')
        else:
            return(self.start)

    def get_booking_time_slot(self):
        hour = self.jamat[:2].strip(':')
        min = self.jamat[-2:].strip(':')

        booking_end = int(min) + 30

        if (booking_end >= 60):
            end_hour = int(hour) + 1
            end_min = booking_end - 60
        else:
            end_hour = hour
            end_min = booking_end

        if len(str(end_hour)) == 1:
            end_hour = '0'+str(end_hour)
        if len(str(end_min)) == 1:
            end_min = '0'+str(end_min)
        return(self.jamat+'-'+str(end_hour)+':'+str(end_min))

    def _style_header(self, ws, row, col):
        if self.header_color:
            ws.cell(row, col).fill = self.header_color
        ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
        ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")

    def add_xl_top_header(self, ws, row, col, do_not_merge=False):
        start_col = col
        ws.cell(row, col).value = self.name
        self._style_header(ws, row, col)
        col += 1
        if self.has_jamat:
            col += 3 # jamat, booking and location columns
        if not do_not_merge:
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=col - 1)
        return col

    def add_xl_header(self, ws, row, col):
        ws.cell(row, col).value = 'Start'
        self._style_header(ws, row, col)
        col += 1
        if self.has_jamat:
            ws.cell(row, col).value = 'Booking_time_slot'
            self._style_header(ws, row, col)
            col += 1
            ws.cell(row, col).value = 'Jamat'
            self._style_header(ws, row, col)
            col += 1
            ws.cell(row, col).value = 'Location'
            self._style_header(ws, row, col)
            col += 1
        return col

    def add_xl_columns(self, ws, row, col):
        ws.cell(row, col).value = self.start
        if self.color_me:
            ws.cell(row, col).fill = self.fill_color
        if self.is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        if self.has_jamat:
            ws.cell(row, col).value = self.booking_start
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            if self.is_juma:
                ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = self.jamat
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            if self.is_juma:
                ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = self.location
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            if self.is_juma:
                ws.cell(row, col).font = Font(bold=True)
            col += 1
        return col


class FajrSalah(Salah):
    def __init__(self, date, start, details, has_jamat=True, fill_color=None, header_color=None):
        super().__init__('Fajr', date, start, details, has_jamat=has_jamat, fill_color=fill_color, header_color=header_color)
        self.sunrise = details['Sunrise']

    def add_xl_top_header(self, ws, row, col):
        start_col = col
        col = super().add_xl_top_header(ws, row, col, do_not_merge=True)
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=col)
        col += 1
        return col

    def add_xl_header(self, ws, row, col):
        col = super().add_xl_header(ws, row, col)
        ws.cell(row, col).value = 'Sunrise'
        self._style_header(ws, row, col)
        col += 1
        return col

    def add_xl_columns(self, ws, row, col):
        col = super().add_xl_columns(ws, row, col)
        ws.cell(row, col).value = self.sunrise
        if self.color_me:
            ws.cell(row, col).fill = self.fill_color
        if self.is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        return col


class JumaSalah(Salah):
    def __init__(self, date, start, details, has_jamat=True, fill_color=None, header_color=None):
        super().__init__('Juma', date, start, details, has_jamat=has_jamat, fill_color=fill_color, header_color=header_color)

    def add_xl_header(self, ws, row, col):
        ws.cell(row, col).value = 'Start'
        self._style_header(ws, row, col)
        col += 1
        ws.cell(row, col).value = 'Booking'
        self._style_header(ws, row, col)
        col += 1
        ws.cell(row, col).value = 'Khutba begines'
        self._style_header(ws, row, col)
        col += 1
        ws.cell(row, col).value = 'Location'
        self._style_header(ws, row, col)
        col += 1
        return col

    def add_xl_columns(self, ws, row, col):
        if self.is_juma:
            ws.cell(row, col).value = "13:05"
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = "13:00"
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = "13:10"
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            ws.cell(row, col).font = Font(bold=True)
            col += 1
            ws.cell(row, col).value = self.location
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
            ws.cell(row, col).font = Font(bold=True)
            col += 1
        else:
            if self.color_me:
                ws.cell(row, col).fill = self.fill_color
                col += 1
                ws.cell(row, col).fill = self.fill_color
                col += 1
                ws.cell(row, col).fill = self.fill_color
                col += 1
                ws.cell(row, col).fill = self.fill_color
                col += 1
            else:
                col += 4
        return col


def salah_org(table):
    fill_grn = PatternFill("solid", fgColor=COLOUR_GREY)
    header_color = PatternFill("solid", fgColor=COLOR_L_GREY)
    for row, (date, day) in enumerate(table['schedule'].items(), start=1):
        fill_color = fill_grn if row % 2 != 0 else None
        Fajr, Sunrise, Juma, Dhuhr, Asr, Maghrib, Isha = day.values()
        day['Fajr'] = FajrSalah(date, Fajr, day, fill_color=fill_color, header_color=header_color)
        day['Dhuhr'] = Salah('Dhuhr', date, Dhuhr, day, has_jamat=False, fill_color=fill_color, header_color=header_color)
        day['Juma'] = JumaSalah(date, Juma, day, has_jamat=True, fill_color=fill_color, header_color=header_color)
        day['Asr'] = Salah('Asr', date, Asr, day, has_jamat=False, fill_color=fill_color, header_color=header_color)
        day['Maghrib'] = Salah('Maghrib', date, Maghrib, day, has_jamat=True, fill_color=fill_color, header_color=header_color)
        day['Isha'] = Salah('Isha', date, Isha, day, has_jamat=True, fill_color=fill_color, header_color=header_color)
        
#    print (table)
    return table


def generate_xl(table, year):
    table = salah_org(table)
    wb = Workbook()
    wb['Sheet'].title = 'CC booking'
    ws = wb['CC booking']
    fill_color = PatternFill("solid", fgColor=COLOUR_GREY)
    header_color = PatternFill("solid", fgColor=COLOR_L_GREY)
    row = 1
    # Add each day rows
    for serial_no, (date, day) in enumerate(table['schedule'].items(), start=1):
        week_day = date.split(' ')[2].strip()
        is_juma = week_day[:3].lower() == 'fri'
        if is_juma:
            fill_color = PatternFill("solid", fgColor=COLOUR_BLUE)
        else:
            fill_color = PatternFill("solid", fgColor=COLOUR_GREY)
        color_me = fill_color if is_juma or serial_no % 2 != 0 else None
        if int(date.split(' ')[1].strip()) == 1:
            # Add top header
            col = 4
            ws.cell(row, 1).value = '{} {}'.format(date.split(' ')[0].strip(), str(year))
            ws.cell(row, 1).font = Font(bold=True, color='FFFFFF')
            ws.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row, 1).fill = header_color
            ws.cell(row, 2).fill = header_color
            ws.cell(row, 3).fill = header_color
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col - 1)
            for salah in next(iter(table['schedule'].values())).values():
                if isinstance(salah, Salah):
                    col = salah.add_xl_top_header(ws, row, col)
            # Add Salah header
            row += 1
            col = 1
            ws.cell(row, col).value = 'Sr. No.'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1
            ws.cell(row, col).value = 'Date'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1
            ws.cell(row, col).value = 'Day'
            ws.cell(row, col).fill = header_color
            ws.cell(row, col).font = Font(bold=True, color='FFFFFF')
            col += 1
            for salah in next(iter(table['schedule'].values())).values():
                if isinstance(salah, Salah):
                    col = salah.add_xl_header(ws, row, col)
            row += 1
        
        col = 1
        ws.cell(row, col).value = serial_no
        serial_no += 1
        if color_me:
            ws.cell(row, col).fill = fill_color
        if is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        ws.cell(row, col).value = date
        if color_me:
            ws.cell(row, col).fill = fill_color
        if is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        ws.cell(row, col).value = date.split(' ')[2]
        if color_me:
            ws.cell(row, col).fill = fill_color
        if is_juma:
            ws.cell(row, col).font = Font(bold=True)
        col += 1
        for salah in day.values():
            if isinstance(salah, Salah):
                col = salah.add_xl_columns(ws, row, col)
        row += 1
    wb.save('Cambourne_salah_timetable_'+year+'.xlsx')


def main():
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('--year', dest='year',
                    default='None', help='Year of salah timetable')
    parser.add_argument('--file', dest='filename',
                    default='None', help='XLS file for timetable')
    args = parser.parse_args()
    table = get_prayer_table_offline(args.year, args.filename)
    generate_xl(table, args.year)


if __name__ == '__main__':
    main()
