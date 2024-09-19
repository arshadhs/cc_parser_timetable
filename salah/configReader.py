import configparser
from datetime import datetime

def get_config(fileName, year):
    config = configparser.ConfigParser()
    config.read(fileName)

    ramadan_start = datetime.strptime(config.get(str(year), 'ramadan_start', fallback=config['DEFAULT']['ramadan_start']), '%Y-%m-%d').date()
    ramadan_end = datetime.strptime(config.get(str(year), 'ramadan_end', fallback=config['DEFAULT']['ramadan_end']), '%Y-%m-%d').date()

    return ramadan_start, ramadan_end
