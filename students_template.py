from openpyxl import Workbook

wb=Workbook()
ws=wb.active
ws.title='Students Details'

headers = [
    "Name",
    "D.O.B",
    "Blood group",
    "Age",
    "Tamil",
    "English",
    "Maths",
    "Physics",
    "Chemistry",
    "Computer science",
    "Total",
    "Average",
    "Rank",
    "Grade",
    "Remarks",
]

ws.append(headers)
wb.save("students_template.xlsx")