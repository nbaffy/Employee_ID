import xlsxwriter
import os


data = [
    {
        'name' : 'Nana Baffour Awuah',
        'position' : 'Intern',
        'contact': '+233-507-170-003',
        'email': 'nbaffy11@gmail.com',
        'address': 'E189, Burma Hills'


    }
]
Workbook = xlsxwriter.Workbook('EmployeeId.xlsx')
sheet = Workbook._add_sheet('Sheet1')

sheet.write(0, 0, '#')
sheet.write(0, 1, 'Name')
sheet.write(0, 2, 'Position')
sheet.write(0, 3, 'Contact')
sheet.write(0, 4, 'Email')
sheet.write(0, 5, 'Address')

for index, entry in enumerate(data):
    sheet.write(index+1, 0, str(index))
    sheet.write(index+1, 1, entry['name'])
    sheet.write(index+1, 2, entry['position'])
    sheet.write(index+1, 3, entry['contact'])
    sheet.write(index+1, 4, entry['email'])
    sheet.write(index+1, 5, entry["address"])
    os.system('start Excel.EXE EmployeeId.xlsx')

