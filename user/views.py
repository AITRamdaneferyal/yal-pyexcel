import xlsxwriter
import json
import requests
from django.http import  HttpResponse
import csv

# make API request and parse JSON automatically
data = requests.get('https://gateway.drsalmi.com/storage/api/files/6305ff598fea21f498f80f54').json()
# ------------------------------------------------------------------------------------------------------------------------------------------------#

# save all data in a single JSON file
file_name = "user_data.json"
with open(file_name, "w") as f:
    json.dump(data, f, indent=4)
print(file_name, "saved successfully!")
#------------------------------------------------------------------------------------------------------------------------------------------------#

#save each entry"user" into a file by rowNumber
#for user in data:
    # iterate over `data` list
   # file_name = f"user_{user['rowNumber']}.json"
   # with open(file_name, "w") as f:
      #  json.dump(user, f, indent=4)
       # print(file_name, "saved successfully!")
#------------------------------------------------------------------------------------------------------------------------------------------------#

# read  all a JSON file
#file_name = "user_data.json"
#with open(file_name) as f:
#    data = json.load(f)
#print(data)
#------------------------------------------------------------------------------------------------------------------------------------------------#
# json data & csv file
#sites = data["records"]
#with open('user.csv', 'w', newline='') as file:
#    writer = csv.writer(file)
#    writer.writerow(["nom", "prenom"])
 #   for elt in data:
  #      print(elt["dataObject"]["customer_family_name"])
   #     print(elt["dataObject"]["customer_first_name"])
    #    temp = elt["dataObject"]["customer_family_name"]
     #   date = elt["dataObject"]["customer_first_name"]
      #  writer.writerow([date, temp])

#------------------------------------------------------------------------------------------------------------------------------------------------#

workbook_name = "excelfile02.xlsx"
workbook = xlsxwriter.Workbook(workbook_name)
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
#Add text size
cell_format.set_font_size(10)
#Add text align
cell_format.set_align('center')
cell_date = workbook.add_format()
cell_date.set_num_format('dd/mm/yyyy hh:mm AM/PM')
worksheet.write(0, 5, 36892.521, cell_date)       # -> 01/01/2001 12:30 AM
widths = [{"name": 30,
        "last_name": 30,
        "age": 10,
        "address": 70
           }]
my_header = ["customer_family_name","customer_first_name","customer_phones","customer_state"]

# Adjust the column width.
#worksheet.set_column(1, 3, 80)  # Width of columns B:D set to 30.
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 30)
worksheet.set_column('C:C', 10)
worksheet.set_column('D:D', 70)


labels = [{"customer_family_name": "nom",
        "customer_first_name": "prénom",
        "customer_phones": "numéro",
        "customer_state": "adresse"
           }]

for index1, entry in enumerate(labels):
       for index2, header in enumerate(my_header):
           worksheet.write(index1, index2, entry[header],cell_format)


#for index1, entry in enumerate(data):
 #      for index2, header in enumerate(my_header):
      #     worksheet.write(index1+1, index2, entry["dataObject"][header])

for elt in data:

        print(elt["dataObject"]["customer_family_name"])
        print(elt["dataObject"]["customer_first_name"])
        print(elt["dataObject"]["customer_phones"])
        print(elt["dataObject"]["customer_state"])
        print(elt["rowNumber"])
        i=elt["rowNumber"]
        nom = elt["dataObject"]["customer_family_name"]
        prénom = elt["dataObject"]["customer_first_name"]
        address = elt["dataObject"]["customer_state"]
        worksheet.write(1+i , 0, nom)
        worksheet.write(1 + i, 1, prénom)
        worksheet.write(1 + i, 3, address)
workbook.close()



#create excel file with statique data
workbook_name = "excelfile01.xlsx"
workbook = xlsxwriter.Workbook(workbook_name)
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
#Add text size
cell_format.set_font_size(10)
#Add text align
cell_format.set_align('center')
cell_date = workbook.add_format()
cell_date.set_num_format('dd/mm/yyyy hh:mm AM/PM')
worksheet.write(0, 5, 36892.521, cell_date)       # -> 01/01/2001 12:30 AM
widths = [{"name": 30,
        "last_name": 30,
        "age": 10,
        "address": 70
           }]
my_header = ["name","last_name","age","address"]

# Adjust the column width.
#worksheet.set_column(1, 3, 80)  # Width of columns B:D set to 30.
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 30)
worksheet.set_column('C:C', 10)
worksheet.set_column('D:D', 70)


labels = [{"name": "nom",
        "last_name": "prénom",
        "age": "age",
        "address": "adresse"
           }]

for index1, entry in enumerate(labels):
       for index2, header in enumerate(my_header):
           worksheet.write(index1, index2, entry[header],cell_format)


data = [
    {
        "name": "aitramdane",
        "last_name": "feryal",
        "age": "24",
        "address": "birkhadem"
    },
    {
        "name": "salmi",
        "last_name": "ahmed",
        "age": "38",
        "address": "babezouar"
    }
]

for index1, entry in enumerate(data):
       for index2, header in enumerate(my_header):
           worksheet.write(index1+1, index2, entry[header])




workbook.close()
print(workbook_name, "saved successfully!")


def home(request):
    return HttpResponse('saved successfully!')