import datetime

def check_format_date(date_string):
    date_format ='%Y-%m-%d'
    try:
        date_obj = datetime.datetime.strptime(date_string, date_format)
    except ValueError:
        print("Incorrect data format, should be YYYY-MM-DD")
