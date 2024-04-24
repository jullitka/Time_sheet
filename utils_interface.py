from datetime import datetime

def get_timestamp(year, month, day):
    my_date = datetime(year, month, day)
    return int(datetime.timestamp(my_date))

def get_timestamp_from_string(string):
    s = string.split('-')
    return get_timestamp(int(s[2]), int(s[1]), int(s[0]))
