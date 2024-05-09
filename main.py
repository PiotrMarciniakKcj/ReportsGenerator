import xlwings as xw
import datetime as datetime
from docxtpl import DocxTemplate


# convert time from float to datetime
def excel_time_to_datetime(excel_time):
    SECONDS_PER_DAY = 86400
    dt = datetime.datetime.utcfromtimestamp(excel_time * SECONDS_PER_DAY)
    return dt.strftime("%H:%M:%S")


# return formated table(without headers) from Excel as 2-dimensional array
def get_and_format_data(start_cell, end_cell, path):
    wb = xw.Book(path, read_only=True, )
    sheet = wb.sheets[0]
    table = sheet.range(start_cell, end_cell).value
    for row in table:
        if row[0] == row[-1]:
            continue
        if row != table[-1]:
            row[0] = excel_time_to_datetime(row[0])
            row[1] = excel_time_to_datetime(row[1])
        else:
            row[0] = ""
        for x in range(2, len(row) - 1):
            row[x] = int(row[x])
        row[len(row) - 1] = ('{:.2f}'.format(round((row[len(row) - 1] * 100), 2)) + "%").replace('.', ',')
    for row in table:
        if row[0] == row[-1]:
            table.remove(row)
    wb.close()
    return table


# get YYYYMMDD date from the name of the Excel file
def get_date(filename):
    #date = filename[filename.find("_") + 1:filename.rfind("_", 0, filename.rfind("_", 0, filename.rfind("_")))]
    date = filename[filename.find("_") + 1:filename.rfind("_", 0, filename.rfind("_"))]
    return date


# copy the contents of the tables from Excel and paste them into the word template
def paste_tables(start_cell, end_cell, file, table):
    for data_row in get_and_format_data(start_cell, end_cell, file):
        table_row = table.add_row().cells
        for x in range(0, len(data_row)):
            table_row[x].text = str(data_row[x])


# paths to Excel files
DR_excel = "102_20230516_090002_DR500_wClass.xlsx"
DP_excel = "102_20230516_140007_DP500_wClass.xlsx"
N_excel = "102_20230515_210000_N200_wClass.xlsx"

# dates of the tests
DR_date = get_date(DR_excel)
DP_date = get_date(DP_excel)
N_date = get_date(N_excel)

# correct order of the tests
order = [DR_date, DP_date, N_date]
order.sort()

# path to template file
template = "template.docx"
output = "output.docx"

doc = DocxTemplate(template)
doc.render({})

tables = doc.tables

# assigning correct table order
DR_detection_table = tables[order.index(DR_date) * 2]
DR_identification_table = tables[order.index(DR_date) * 2 + 1]
DP_detection_table = tables[order.index(DP_date) * 2]
DP_identification_table = tables[order.index(DP_date) * 2 + 1]
N_detection_table = tables[order.index(N_date) * 2]
N_identification_table = tables[order.index(N_date) * 2 + 1]

# pasting the tables
paste_tables('I2', 'N9', DR_excel, DR_detection_table)
paste_tables('R2', 'X9', DR_excel, DR_identification_table)
paste_tables('I2', 'N18', N_excel, N_detection_table)
paste_tables('R2', 'X18', N_excel, N_identification_table)
paste_tables('I2', 'N8', DP_excel, DP_detection_table)
paste_tables('R2', 'X8', DP_excel, DP_identification_table)

doc.save(output)
