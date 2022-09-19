import xlsxwriter

workbook = xlsxwriter.Workbook('excelfile1.xlsx')
worksheet = workbook.add_worksheet()
my_header = ["name","last_name","age","address"]

for col_num, header in enumerate(my_header):
    worksheet.write(0, col_num, header)

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