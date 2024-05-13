import datetime as datetime
from docxtpl import DocxTemplate
from openpyxl import load_workbook


# convert float to string value with '%'
def float_to_percent(value):
    return ('{:.2f}'.format(round((value * 100), 2)) + "%").replace('.', ',')


# return index of the last row of the table
def enumerate_sheet(worksheet, max_val):
    for i, i_row in enumerate(worksheet):
        if i_row[max_val].value == "Suma:":
            return str(i+1)


# return formatted tables from an Excel file
def get_formatted_data(path):
    wb = load_workbook(path)
    sheet = wb.worksheets[0]
    tables_from_excel = [format_data("I2:N" + enumerate_sheet(sheet, 9), path),
                         format_data("R2:X" + enumerate_sheet(sheet, 18), path)]
    return tables_from_excel


# return formatted table(without headers) from Excel as a 2-dimensional array from an index input
def format_data(indexes, path):
    wb = load_workbook(path)
    sheet = wb.worksheets[0]
    input_table = sheet[indexes]
    new_table = []
    for row in input_table:
        temp_row = []
        if row[0].value == row[-1].value:
            # skip empty rows
            continue
        for cell in row:
            cell = cell.value
            if cell is None:
                cell = ''
            if type(cell) is datetime.time:
                cell = cell.strftime("%H:%M:%S")
            temp_row.append(cell)
        row = temp_row
        row[len(row) - 1] = float_to_percent(float(row[len(row) - 1]))
        new_table.append(row)
    print(new_table[-1])
    return new_table


# get YYYYMMDD date from the name of the Excel file
def get_date(filename):
    date = filename[filename.find("_") + 1:filename.rfind("_", 0, filename.rfind("_"))]
    return date


# copy the contents of the tables from Excel and paste them into the word template
def paste_tables(formated_data, table_name, is_summary=False):
    for data_row in formated_data:
        if is_summary:
            row = table_name.rows[1].cells
            for x in range(0, 4):
                row[x].text = str(data_row[x])
        else:
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
formatted_tables = get_formatted_data(DR_excel)
formatted_tables.extend(get_formatted_data(DP_excel))
formatted_tables.extend(get_formatted_data(N_excel))

print(formatted_tables)

N_sum = 0
em_sum = 0
ef_sum = 0

Nid_sum = 0
Kok_sum = 0
rejected_sum = 0

for table in formatted_tables:
    if formatted_tables.index(table) % 2 == 0:
        N_sum += table[-1][2]
        em_sum += table[-1][3]
        ef_sum += table[-1][4]
    else:
        Nid_sum += table[-1][2]
        Kok_sum += table[-1][3]
        rejected_sum += table[-1][5]

d = float_to_percent((N_sum - em_sum - ef_sum) / N_sum)
r = float_to_percent(Kok_sum / Nid_sum)

# 2-dimensional array with just one row
summary_detection = [[N_sum, em_sum, ef_sum, d]]
summary_identification = [[Nid_sum, Kok_sum, r, rejected_sum]]

# pasting the tables
paste_tables(formatted_tables[0], DR_detection_table)
paste_tables(formatted_tables[1], DR_identification_table)
paste_tables(formatted_tables[2], N_detection_table)
paste_tables(formatted_tables[3], N_identification_table)
paste_tables(formatted_tables[4], DP_detection_table)
paste_tables(formatted_tables[5], DP_identification_table)
paste_tables(summary_detection, summary_detection_table, True)
paste_tables(summary_identification, summary_identification_table, True)

doc.save(output)
