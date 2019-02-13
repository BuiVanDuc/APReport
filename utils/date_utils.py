import datetime

def parse_date_from_string(date_string):
    date_format ='%Y-%m-%d'
    try:
        date_obj = datetime.datetime.strptime(date_string, date_format)
        return date_obj
    except ValueError:
        print("Incorrect data format, should be YYYY-MM-DD")
        return False