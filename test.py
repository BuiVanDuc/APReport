from datetime import datetime

list_data = [{"a":2, "b":2,"c":"ccc"},{"a":2, "b":1,"c":"xx"}]
data = dict()

# for i in range(1,len(list_data)):
#     for key, val in list_data[0].items():
#         if isinstance(list_data[0][key], str):
#             data[key] = ""
#         else:
#             sum_number = sum(list_data[0][key] + list_data[i][key])
#             data[key] = sum_number[0][key]
date=datetime.today()
print date

year=date.strftime("%Y")

print year