import xlwings as xw
import datetime as datetime
from docxtpl import DocxTemplate


# convert time from float to datetime
def excel_time_to_datetime(excel_time):
    SECONDS_PER_DAY = 86400
    dt = datetime.datetime.utcfromtimestamp(excel_time * SECONDS_PER_DAY)
    return dt.strftime("%H:%M:%S")


# convert float to string value with '%'
def float_to_percent(value):
    return ('{:.2f}'.format(round((value * 100), 2)) + "%").replace('.', ',')


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
        row[len(row) - 1] = float_to_percent(row[len(row) - 1])
    for row in table:
        if row[0] == row[-1]:
            # delete empty rows
            table.remove(row)
    wb.close()
    return table


# get YYYYMMDD date from the name of the Excel file
def get_date(filename):
    date = filename[filename.find("_") + 1:filename.rfind("_", 0, filename.rfind("_"))]
    return date


# copy the contents of the tables from Excel and paste them into the word template
def paste_tables(formated_data, table_name):
    for data_row in formated_data:
        table_row = table_name.add_row().cells
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
summary_detection_table = tables[6]
summary_identification_table = tables[7]

# getting the tables from the Excel files
formated_tables = [
    get_and_format_data('I2', 'N9', DR_excel),
    get_and_format_data('R2', 'X9', DR_excel),
    get_and_format_data('I2', 'N18', N_excel),
    get_and_format_data('R2', 'X18', N_excel),
    get_and_format_data('I2', 'N8', DP_excel),
    get_and_format_data('R2', 'X8', DP_excel)
]

N_sum = 0
em_sum = 0
ef_sum = 0

Nid_sum = 0
Kok_sum = 0
rejected_sum = 0

for table in formated_tables:
    if formated_tables.index(table) % 2 == 0:
        N_sum += table[-1][2]
        em_sum += table[-1][3]
        ef_sum += table[-1][4]
    else:
        Nid_sum += table[-1][2]
        Kok_sum += table[-1][3]
        rejected_sum += table[-1][5]

d = float_to_percent((N_sum - em_sum - ef_sum) / N_sum)
r = float_to_percent(Kok_sum / Nid_sum)

summary_detection = [N_sum, em_sum, ef_sum, d]
summary_identification = [Nid_sum, Kok_sum, r, rejected_sum]


# pasting the tables
paste_tables(formated_tables[0], DR_detection_table)
paste_tables(formated_tables[1], DR_identification_table)
paste_tables(formated_tables[2], N_detection_table)
paste_tables(formated_tables[3], N_identification_table)
paste_tables(formated_tables[4], DP_detection_table)
paste_tables(formated_tables[5], DP_identification_table)
# doesn't work :(
# paste_tables(summary_detection, summary_detection_table)
# paste_tables(summary_identification, summary_identification_table)

doc.save(output)
