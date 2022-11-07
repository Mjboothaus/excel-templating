# https://medium.com/geekculture/automate-your-excel-file-generation-with-python-42552dabd654

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

REGULAR_SIZE = 11
REGULAR_FONT = 'Cambria'

def default_format(workbook):
    default = workbook.add_format({
        'font_name':REGULAR_FONT,
        'font_size':REGULAR_SIZE,
        'valign':'top',
    })
    return default

def text_box_wrap_format(workbook):
    text_box_wrap = workbook.add_format({
        'font_name':REGULAR_FONT,
        'font_size':REGULAR_SIZE,
        'align':'justify',
        'valign':'vcenter',
        'border':True,
        'text_wrap':True
    })
    return text_box_wrap

def text_box_center_wrap_format(workbook):
    text_box_center_wrap = workbook.add_format({
        'font_name':REGULAR_FONT,
        'font_size':REGULAR_SIZE,
        'align':'center',
        'valign':'vcenter',
        'border':True,
        'text_wrap':True
    })
    return text_box_center_wrap

workbook = xlsxwriter.Workbook('grading.xlsx')
worksheet = workbook.add_worksheet('Assignment 1')

worksheet.set_column('B:B', 15)

row = 0
col = 0

worksheet.merge_range(row, col, row+2, col, "No.", text_box_center_wrap_format(workbook))
worksheet.merge_range(row, col+1, row+2, col+1, "Name", text_box_center_wrap_format(workbook))

indicators = [
    ['Correctness', 40],
    ['Doccumentation', 40],
    ['Demo', 20]
]
tmp_col = col+2+len(indicators)-1
worksheet.merge_range(
    row, col+2, row, tmp_col, "Grading Indicators", 
    text_box_center_wrap_format(workbook)
)
for idx, indicator in enumerate(indicators):
    worksheet.write(row+1, col+2+idx, idx+1, text_box_center_wrap_format(workbook))
    worksheet.write(row+2, col+2+idx, indicator[1], text_box_center_wrap_format(workbook))

worksheet.merge_range(row, tmp_col+1, row+2, tmp_col+1, "Total", text_box_center_wrap_format(workbook))

row += 3

students_grade = {
    "Nanda" : [40, 40, 20],
    "Ryaas" : [28, 32, 15],
    "Absar" : [38, 32, 5]
}

for idx, (name, grades) in enumerate(sorted(students_grade.items())):
    worksheet.write(row, col, idx+1, text_box_center_wrap_format(workbook))
    worksheet.write(row, col+1, name, text_box_wrap_format(workbook))
    st_cell_grade = xl_rowcol_to_cell(row, col+2)
    tmp_col = col+2
    for grade in grades:
        worksheet.write(row, tmp_col, grade, text_box_center_wrap_format(workbook))
        tmp_col += 1
    end_cell_grade = xl_rowcol_to_cell(row, tmp_col-1)
    worksheet.write(
        row, tmp_col, "=SUM({}:{})".format(st_cell_grade, end_cell_grade),
        text_box_center_wrap_format(workbook)
    )
    row += 1

row += 1
worksheet.write(row, col, "Indicators:", default_format(workbook))
row += 1
for idx, indicator in enumerate(indicators):
    placeholder = "{}:{}".format(idx+1, indicator[0])
    worksheet.write(row, col, placeholder, default_format(workbook))
    row += 1

workbook.close()
