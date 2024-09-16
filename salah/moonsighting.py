import json
from typing import OrderedDict
import requests
import xmltodict
from openpyxl import load_workbook, Workbook

def get_prayer_table(year):
    url = 'https://www.moonsighting.com/praytable.php'
    parameters = {'year': str(year), 'tz' : 'Europe/London', 'lat': '52.2178,', 'lon': '0.0662', 'method': '0', 'both': 'false', 'time': '0'}
    response = requests.get(url, params=parameters)
    if response.status_code != 200:
        raise Exception(f"Failed to get prayer time table from {url}")
    start = response.text.find('<div')
    end = response.text.rfind('</div>')
    xml_text = response.text[start:end+len('</div>')]
    elements = xmltodict.parse(xml_text)
    header = elements['div']['table']['thead']['tr']['th']
    header[0] = 'Date'
    data = OrderedDict()
    data = {'header': header}
    schedule = OrderedDict()
    for day in elements['div']['table']['tbody']['tr']:
        date, Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = day["td"]
        schedule[date] = OrderedDict(Fajr=Fajr, Sunrise=Sunrise, Dhuhr=Dhuhr, Juma=Dhuhr, Asr=Asr, Maghrib=Maghrib, Isha=Isha)
    data['schedule'] = schedule
    return data

def _get_sheet_from_hdr(wb, headers):
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows(min_row=1, max_col=7, max_row=1):
            header = []

            for cell in row:
                # print (cell.value)
                # print (type(cell.value))
                header.append(str(cell.value))
            break               

        if set(header).issuperset(headers):
            print(f"Found xls '{sheet_name}'")
            return sheet
        else:
            print (set(header).difference(headers))
            print (set(headers).difference(header))

    print('Failed to find the xls with timing information')
    return None

def get_donation_sheet(filename):
    iwb = load_workbook(filename, read_only=True)
    return _get_sheet_from_hdr(iwb, {"2025", 'Fajr', 'Sunrise', 'Dhuhr', 'Asr(H)', 'Maghrib', 'Isha'})


def get_prayer_table_offline(year, filename):
    sheet = get_donation_sheet(filename)

#    print ("sheet: ", type(sheet))
    schedule = OrderedDict()

#    header[0] = 'Date'
    data = OrderedDict()
#    data = {'header': header}

    if sheet:
        for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):

            date, Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = row[:7]
#            print (row)
            schedule[date] = OrderedDict(Fajr=str(Fajr), Sunrise=str(Sunrise), Dhuhr=str(Dhuhr), Juma=str(Dhuhr), Asr=str(Asr), Maghrib=str(Maghrib), Isha=str(Isha))

#    for key, value in (schedule).items():
        # print ("key: ", key)
        # print ("value: ", type(schedule))
        # print ("value: ", type(value))
        # print ("value: ", value)
        
        # for k, v in (value).items():
                # print ("key: ", k)
                # print ("value: ", type(v))
                # print ("value: ", v)

    data['schedule'] = schedule
    return data

    # url = 'https://www.moonsighting.com/praytable.php'
    # parameters = {'year': str(year), 'tz' : 'Europe/London', 'lat': '52.2178,', 'lon': '0.0662', 'method': '0', 'both': 'false', 'time': '0'}
    # response = requests.get(url, params=parameters)
    # if response.status_code != 200:
        # raise Exception(f"Failed to get prayer time table from {url}")
    # start = response.text.find('<div')
    # end = response.text.rfind('</div>')
    # xml_text = response.text[start:end+len('</div>')]
    # elements = xmltodict.parse(xml_text)
    # header = elements['div']['table']['thead']['tr']['th']
    
    
    
    # header[0] = 'Date'
    # data = OrderedDict()
    # data = {'header': header}
    # schedule = OrderedDict()
    # for day in elements['div']['table']['tbody']['tr']:
        # date, Fajr, Sunrise, Dhuhr, Asr, Maghrib, Isha = day["td"]
        # schedule[date] = OrderedDict(Fajr=Fajr, Sunrise=Sunrise, Dhuhr=Dhuhr, Juma=Dhuhr, Asr=Asr, Maghrib=Maghrib, Isha=Isha)
    # data['schedule'] = schedule
    # return data
