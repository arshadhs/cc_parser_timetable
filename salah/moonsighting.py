import json
from typing import OrderedDict
import requests
import xmltodict


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
