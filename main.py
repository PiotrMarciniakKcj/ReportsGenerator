import types

import xlwings as xw
import datetime as datetime
from docxtpl import DocxTemplate


# converts time from float to datetime
def excel_time_to_datetime(excel_time):
    SECONDS_PER_DAY = 86400
    dt = datetime.datetime.utcfromtimestamp(excel_time * SECONDS_PER_DAY)
    return dt.strftime("%H:%M:%S")


# returns formated table(without headers) from Excel as 2-dimensional array
def get_and_format_data(start_cell, end_cell, path):
    wb = xw.Book(path, read_only=True, )
    sheet = wb.sheets[0]
    table = sheet.range(start_cell, end_cell).value
    for row in table:
        if row != table[-1]:
            row[0] = excel_time_to_datetime(row[0])
            row[1] = excel_time_to_datetime(row[1])
        else:
            row[0] = ""
        for x in range(2, len(row) - 1):
            row[x] = int(row[x])
        row[len(row) - 1] = ('{:.2f}'.format(round((row[len(row) - 1] * 100), 2)) + "%").replace('.', ',')
    wb.close()
    return table


# get YYYYMMDD date from the name of the Excel file
def get_date_from_filename(filename):
    date = filename[filename.find("_") + 1:filename.rfind("_", 0, filename.rfind("_", 0, filename.rfind("_")))]
    return date


# path = "C:\\Users\\User\Desktop\\102_20230516_090002_DR500_wClass.xlsx"

# paths to Excel files
DR_excel = "102_20230516_090002_DR500_wClass.xlsx"
DP_excel = "102_20230516_140007_DP500_wClass.xlsx"
N_excel = "102_20230515_210000_N200_wClass.xlsx"

DR_date = get_date_from_filename(DR_excel)
print(DR_date)

# path to template file
template = "template.docx"
output = "output.docx"

print(get_and_format_data('I2', 'N9', DR_excel))
print(get_and_format_data('R2', 'X9', DR_excel))

doc = DocxTemplate(template)
doc.render({})

tables = doc.tables

DR_detection_table = tables[0]
DR_identification_table = tables[1]
DP_detection_table = tables[2]
DP_identification_table = tables[3]
N_detection_table = tables[4]
N_identification_table = tables[5]

for data_row in get_and_format_data('I2', 'N9', DR_excel):
    table_row = DR_detection_table.add_row().cells
    for x in range(0, len(data_row)):
        table_row[x].text = str(data_row[x])

doc.save(output)
