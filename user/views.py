import xlsxwriter
import json
import requests
from django.http import HttpResponse
import csv
# classse user
from typing import List
from typing import Any
from dataclasses import dataclass
from datetime import datetime

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
    save_path = "C:/Users/HB/Desktop/json file/"
    with open("  ../../user_data.json", "w") as f:
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

    # create the HttpResponse object ...
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = "attachment; filename=excelfile02.xlsx"
    ############# partie03 json file and xlsx file #############
    workbook_name = "excelfile02.xlsx"

    workbook = xlsxwriter.Workbook(response, {'in_memory': True})
    #workbook = xlsxwriter.Workbook(workbook_name)  # creation du fichier
    ################################feuille 01 #############################
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
       # print(elt["dataObject"]["customer_family_name"])
        #print(elt["dataObject"]["customer_first_name"])
        #print(elt["dataObject"]["customer_phones"])
       # print(elt["dataObject"]["customer_state"])
       # print(elt["rowNumber"])
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
        worksheet.write(i, 2, ', '.join(use.num), cell_format_center)
        worksheet.write(i, 3, use.state, cell_format_center)


    ################################feuille 02 #############################
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
    ################################feuille 03 #############################
    worksheet3 = workbook.add_worksheet("calculate_cells")
    # Add a number format for cells with money.
    money_format = workbook.add_format({'num_format': '#,##0$'})
    # Add an Excel date format.
    date_format = workbook.add_format({'num_format': 'dd mmmm yy'})
    # Adjust the column width.
    worksheet3.set_column(1, 1, 15)
    # Some data we want to write to the worksheet.
    expenses = (
        ['Rent', '2013-01-13', 1000],
        ['Gas', '2023-06-14', 100],
        ['Rent', '2016-02-16', 300],
        ['Gym', '2022-01-20', 50],
    )

    # Start from the first cell below the headers.
    row = 1
    col = 0

    for item, date_str, cost in (expenses):
        # Convert the date string into a datetime object.
        date = datetime.strptime(date_str, "%Y-%m-%d")

        worksheet3.write_string(row, col, item)
        worksheet3.write_datetime(row, col + 1, date, date_format)
        worksheet3.write_number(row, col + 2, cost, money_format)
        worksheet3.write_formula(row, col + 3, '=IF(C:C>100,"Yes", "No")')

        row += 1
    worksheet3.write_formula("K2", 'Rent')
    # Write a total using a formula sum.
    worksheet3.write(row, 0, 'Total')
    worksheet3.write(row, 2, '=SUM(B1:B4)',money_format)
    # filter .
    worksheet3.write('F2', '=FILTER(A1:D6,A1:A6=k2)')
    ################################feuille 04 #############################
    worksheet4 = workbook.add_worksheet("data format")
    # Adjust the column width.
    worksheet4.set_column('A:B', 30)

    cell_format1 = workbook.add_format()
    cell_format2 = workbook.add_format()
    cell_format01 = workbook.add_format()
    cell_format02 = workbook.add_format()
    cell_format03 = workbook.add_format()
    cell_format04 = workbook.add_format()
    cell_format05 = workbook.add_format()
    cell_format06 = workbook.add_format()
    cell_format07 = workbook.add_format()
    cell_format08 = workbook.add_format()
    cell_format09 = workbook.add_format()
    cell_format10 = workbook.add_format()
    cell_format11 = workbook.add_format()


    cell_format1.set_num_format('dddd mmm yyyy')  # Format string.
    worksheet4.write(1, 1, 3.1415926, cell_format1)
    cell_format2.set_num_format(0x0F)  # Format index.
    worksheet4.write(1, 2, 3.1415926, cell_format2)

    cell_format01.set_num_format('0.000')
    worksheet4.write(1, 0, 3.1415926, cell_format01)  # -> 3.142

    cell_format02.set_num_format('#,##0')
    worksheet4.write(2, 0, 1234.56, cell_format02)  # -> 1,235

    cell_format03.set_num_format('#,##0.00')
    worksheet4.write(3, 0, 1234.56, cell_format03)  # -> 1,234.56

    cell_format04.set_num_format('0.00')
    worksheet4.write(4, 0, 49.99, cell_format04)  # -> 49.99

    cell_format05.set_num_format('mm/dd/yy')
    worksheet4.write(5, 0, 36892.521, cell_format05)  # -> 01/01/01

    cell_format06.set_num_format('mmm d yyyy')
    worksheet4.write(6, 0, 36892.521, cell_format06)  # -> Jan 1 2001

    cell_format07.set_num_format('d mmmm yyyy')
    worksheet4.write(7, 0, 36892.521, cell_format07)  # -> 1 January 2001

    cell_format08.set_num_format('dd/mm/yyyy hh:mm AM/PM')
    worksheet4.write(8, 0, 36892.521, cell_format08)  # -> 01/01/2001 12:30 AM

    cell_format09.set_num_format('0 "dollar and" .00 "cents"')
    worksheet4.write(9, 0, 1.87, cell_format09)  # -> 1 dollar and .87 cents

    # Conditional numerical formatting.
    cell_format10.set_num_format('[Green]General;[Red]-General;General')
    worksheet4.write(10, 0, 123, cell_format10)  # > 0 Green
    worksheet4.write(11, 0, -45, cell_format10)  # < 0 Red
    worksheet4.write(12, 0, 0, cell_format10)  # = 0 Default color

    # Zip code.
    cell_format11.set_num_format('00000')
    worksheet4.write(13, 0, 1209, cell_format11)
    ################################feuille 05 #############################
    worksheet5 = workbook.add_worksheet("date format")
    # Adjust the column width.
    worksheet5.set_column('A:F', 30)
    # Write the column headers.
    worksheet5.write('A1', 'Formatted date')
    worksheet5.write('B1', 'Format')
    # Create a datetime object to use in the examples.
    date_time = datetime.strptime('2013-01-23 12:30:05.123',
                                  '%Y-%m-%d %H:%M:%S.%f')
    # Examples date and time formats.
    date_formats = (
        'dd/mm/yy',
        'mm/dd/yy',
        'dd m yy',
        'd mm yy',
        'd mmm yy',
        'd mmmm yy',
        'd mmmm yyy',
        'd mmmm yyyy',
        'dd/mm/yy hh:mm',
        'dd/mm/yy hh:mm:ss',
        'dd/mm/yy hh:mm:ss.000',
        'hh:mm',
        'hh:mm:ss',
        'hh:mm:ss.000',
    )
    # Start from first row after headers.
    row = 1
    # Write the same date and time using each of the above formats.
    for date_format_str in date_formats:
        # Create a format for the date or time.
        date_format = workbook.add_format({'num_format': date_format_str,
                                           'align': 'left'})
        # Write the same date using different formats.
        worksheet5.write_datetime(row, 0, date_time, date_format)
        # Also write the format string for comparison.
        worksheet5.write_string(row, 1, date_format_str)
        row += 1
        ################################feuille 06 #############################
    worksheet6 = workbook.add_worksheet("validate liste")
    # Adjust the column width.
    worksheet6.set_column('A:F', 30)
    worksheet6.write('C1',"liste")
    # field list (liste deroulante)
    validate_liste = {'validate': 'list',
                      'source': ['open', 'high', 'close']}
    worksheet6.data_validation('C2:C12', validate_liste)
    ################################feuille 07 #############################
    worksheet7 = workbook.add_worksheet("Merged Cells")


    merge_format = workbook.add_format({'align': 'center'})
    worksheet7.merge_range('B3:D4', 'Merged Cells', merge_format)


    workbook.close()
    print(workbook_name, "saved successfully!")


    return response

