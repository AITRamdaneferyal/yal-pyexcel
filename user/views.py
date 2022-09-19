import xlsxwriter

workbook = xlsxwriter.Workbook('excelfile1.xlsx')
worksheet = workbook.add_worksheet()
my_header = ["name","last_name","age","address"]


labels = [{"name": "nom",
        "last_name": "prénom",
        "age": "age",
        "address": "adresse"
           }]

for index1, entry in enumerate(labels):
       for index2, header in enumerate(my_header):
           worksheet.write(index1, index2, entry[header])


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