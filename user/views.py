import xlsxwriter
import json
import requests
from django.http import HttpResponse
import csv
# classse user
from typing import List
from typing import Any
from dataclasses import dataclass
import io


@dataclass
class user:
    family_name: str
    first_name: str
    num: List[str]
    state: str


# django http response
def home(request):
    ############# partie01 json file #############
    # make API request and parse JSON automatically
    data = requests.get('https://gateway.drsalmi.com/storage/api/files/6305ff598fea21f498f80f54').json()
    # ------------------------------------------------------------------------------------------------------------------------------------------------#

    # save all data in a single JSON file
    file_name = "user_data.json"
    with open(file_name, "w") as f:
        json.dump(data, f, indent=4)
    print(file_name, "saved successfully!")
    # ------------------------------------------------------------------------------------------------------------------------------------------------#

    # save each entry"user" into a file by rowNumber
    # for user in data:
    # iterate over `data` list
    # file_name = f"user_{user['rowNumber']}.json"
    # with open(file_name, "w") as f:
    #  json.dump(user, f, indent=4)
    # print(file_name, "saved successfully!")
    # ------------------------------------------------------------------------------------------------------------------------------------------------#

    # read  all a JSON file
    # file_name = "user_data.json"
    # with open(file_name) as f:
    #    data = json.load(f)
    # print(data)
    # ------------------------------------------------------------------------------------------------------------------------------------------------#

    ############# partie02 json file and csv file #############
    # json data & csv file
    # sites = data["records"]
    # with open('user.csv', 'w', newline='') as file:
    #    writer = csv.writer(file)
    #    writer.writerow(["nom", "prenom"])
    #   for elt in data:
    #      print(elt["dataObject"]["customer_family_name"])
    #     print(elt["dataObject"]["customer_first_name"])
    #    temp = elt["dataObject"]["customer_family_name"]
    #   date = elt["dataObject"]["customer_first_name"]
    #  writer.writerow([date, temp])
    # ------------------------------------------------------------------------------------------------------------------------------------------------#

    ############# partie03 json file and xlsx file #############
    workbook_name = "excelfile02.xlsx"
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    #workbook = xlsxwriter.Workbook(workbook_name)  # creation du fichier
    worksheet = workbook.add_worksheet("feuille1")  # creation de la feuille1
    # Add a bold format to use to highlight cells.
    cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    cell_format_center = workbook.add_format()
    cell_format_center.set_align('center')

    # Add text size
    cell_format.set_font_size(10)
    # Add text align
    cell_format.set_align('center')
    # format_left_to_right
    format_left_to_right = workbook.add_format()
    format_left_to_right.set_reading_order(1)
    # format_right_to_left
    format_right_to_left = workbook.add_format()
    format_right_to_left.set_reading_order(2)
    # sheet orientation
    # worksheet.right_to_left()

    # add date format
    cell_date = workbook.add_format()
    cell_date.set_num_format('dd/mm/yyyy hh:mm AM/PM')
    worksheet.write(0, 5, 36892.521, cell_date)  # -> 01/01/2001 12:30 AM

    widths = [{"name": 30,
               "last_name": 30,
               "age": 10,
               "address": 70
               }]
    # Adjust the column width.
    # worksheet.set_column(1, 3, 80)  # Width of columns B:D set to 30.
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 40)
    worksheet.set_column('D:D', 70)
    worksheet.set_column(4, 7, 90)

    my_header = ["customer_family_name", "customer_first_name", "customer_phones", "customer_state"]
    labels = [{"customer_family_name": "nom",
               "customer_first_name": "prénom",
               "customer_phones": "numéro",
               "customer_state": "adresse"
               }]
    # add header with cell_format
    for index1, entry in enumerate(labels):
        for index2, header in enumerate(my_header):
            worksheet.write(index1, index2, entry[header], cell_format)

    # add data
    for elt in data:
        print(elt["dataObject"]["customer_family_name"])
        print(elt["dataObject"]["customer_first_name"])
        print(elt["dataObject"]["customer_phones"])
        print(elt["dataObject"]["customer_state"])
        print(elt["rowNumber"])
        i = elt["rowNumber"]
        use = user(family_name=elt["dataObject"]["customer_family_name"],
                   first_name=elt["dataObject"]["customer_first_name"], state=elt["dataObject"]["customer_state"],
                   num=elt["dataObject"]["customer_phones"])
        # nom = user(family_name =elt["dataObject"]["customer_family_name"])
        # prénom = user(first_name =elt["dataObject"]["customer_first_name"])
        # address = user(state =elt["dataObject"]["customer_state"])
        print((use.num))
        worksheet.write(i, 0, use.family_name, cell_format_center)
        worksheet.write(i, 1, use.first_name, cell_format_center)
        worksheet.write(i, 3, use.state, cell_format_center)
        row = i
        col = 4
        for module in use.num:
            str1 = module[0]
            if str1 is None:  continue
            worksheet.write_row(row, col, module)
            row += 1

    # ajouter deuxiéme feuille
    worksheet2 = workbook.add_worksheet("feuille2")  # creation de la feuille
    # Add a bold format to use to highlight cells.
    cell_format1 = workbook.add_format({'bold': True, 'font_color': '#33B3A2'})
    cell_format1.set_font_name('Times New Roman')
    cell_format1.set_font_size(20)
    cell_format1.set_bg_color('red')
    cell_format2 = workbook.add_format({'bold': True})
    cell_format3 = workbook.add_format({'bold': True, 'font_color': 'red'})
    # Add text size
    cell_format.set_font_size(10)
    # Add text align
    cell_format.set_align('center')
    # format_right_to_left
    format_right_to_left = workbook.add_format()
    format_right_to_left.set_reading_order(2)

    my_header = ["name", "last_name", "age", "address"]

    # Adjust the column width.
    worksheet2.set_column('A:A', 50)
    worksheet2.set_column('B:B', 50)
    worksheet2.set_column('C:C', 50)
    worksheet2.set_column('D:D', 50)
    worksheet2.set_column('E:E', 70)
    # Adjust the row width.
    worksheet2.set_row(5, 70)

    labels = [{
        "name": "nom",
        "last_name": "prénom",
        "age": "age",
        "address": "adresse"
    }]
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
    worksheet2.write(0, 0 + 1, "ID", cell_format1)
    for index1, entry in enumerate(labels):
        for index2, header in enumerate(my_header):
            worksheet2.write(index1, index2 + 1, entry[header], cell_format1)

    for index1, entry in enumerate(data):
        for index2, header in enumerate(my_header):
            worksheet2.write(index1 + 1, index2 + 1, entry[header], format_right_to_left)
            worksheet2.write(index1 + 1, 0, index1, cell_format2)

    workbook.close()

    print(workbook_name, "saved successfully!")

    # return HttpResponse('saved successfully!')
    #response = HttpResponse(content_type='application/ms-excel')
    #response['Content-Disposition'] = 'attachment; filename="excelfile02.xls"'
    output.seek(0)

    response = HttpResponse(output.read(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=test.xlsx"

    output.close()

    return response

