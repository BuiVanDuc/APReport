from datetime import datetime
from dateutil import parser

def parse_date_from_string(date_string):
    try:
        date = parser.parse(date_string)
        return date
    except ValueError:
        print("Incorrect data format, should be YYYY-MM-DD")
        return False


def convert_datetime_to_string(date_time, type_format):
    if type_format == 1:
        date_time_str = date_time.strftime("%Y_%m_%d_T%H_%M_%S")
        return date_time_str
    elif type_format == 2:
        date_time_str = date_time.strftime("%Y_%m_%d")
        return date_time_str

def validate(date_text, date_format):
    try:
        date_obj = datetime.strptime(date_text, date_format)
        return date_obj
    except ValueError:
        print ("Incorrect data format, should be YYYY-MM-DD")

