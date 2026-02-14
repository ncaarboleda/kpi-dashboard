from pathlib import Path
from openpyxl import Workbook
from datetime import date
from openpyxl.styles import Font, PatternFill, Side, Border, Alignment
import openpyxl
import os
import math

today = date.today()
current_year = today.year
current_month = today.month
excel_name = f'{current_year}-{current_month}'

green_fill = PatternFill(start_color='009E4D', end_color='009E4D', fill_type='solid')
center_alignment = Alignment(horizontal='center')
bold_font = Font(bold=True, size=12)

raw_directory = Path.cwd() /"resources"/"raw"
current_directory = Path.cwd()

excel_files_xlsx = list(raw_directory.glob('*.xlsx'))
excel_files_xls = list(raw_directory.glob('*.xls'))
all_excel_files = excel_files_xlsx + excel_files_xls

def exec():
    print(f'[LOG] Total files: {all_excel_files}')

    file_path_month = current_directory / f'{excel_name}.xlsx'
    if os.path.isfile(file_path_month):
        print(f"[LOG] The file '{file_path_month}' exists.")

    else:
        print(f"[LOG] The file '{file_path_month}' does not exist.")
        createWb()

    process(all_excel_files)

def process(all_excel_files):
    print(f'[LOG] Starting to process the files: {all_excel_files}')

    for file_path in all_excel_files:
        print(f"[LOG] Processing file: {file_path}")
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook['Technical']

        name = sheet.cell(row=3, column=3).value
        id = sheet.cell(row=4, column=3).value
        score = sheet.cell(row=3, column=5).value
        date = sheet.cell(row=6, column=3).value

        workbookMonth = openpyxl.load_workbook(f'{excel_name}.xlsx')
        sheetQA = workbook['Quality Audit']


def createWb():
    wb = Workbook()
    ws = wb.active

    ws.title = "Dashboard"
    createWorksheet(ws, 1, "Employee ID")
    createWorksheet(ws, 2, "Employee Name")
    createWorksheet(ws, 3, "Quality Audit")
    createWorksheet(ws, 4, "Customer Satisfaction")
    createWorksheet(ws, 5, "MLL Hours")
    createWorksheet(ws, 6, "Total Scores")

    ws2 = wb.create_sheet('Quality Audit')
    createWorksheet(ws2, 1, "Employee ID")
    createWorksheet(ws2, 2, "Name")
    createWorksheet(ws2, 3, "W1")
    createWorksheet(ws2, 4, "W2")
    createWorksheet(ws2, 5, "W3")
    createWorksheet(ws2, 6, "W4")
    createWorksheet(ws2, 7, "W5")
    createWorksheet(ws2, 8, "W6")
    createWorksheet(ws2, 9, "W7")

    ws3 = wb.create_sheet('Customer Satisfaction')
    createWorksheet(ws3, 1, "Employee ID")
    createWorksheet(ws3, 2, "Name")
    createWorksheet(ws3, 3, "W1")
    createWorksheet(ws3, 4, "W2")
    createWorksheet(ws3, 5, "W3")
    createWorksheet(ws3, 6, "W4")
    createWorksheet(ws3, 7, "W5")
    createWorksheet(ws3, 8, "W6")
    createWorksheet(ws3, 9, "W7")


    wb.save(f'{excel_name}.xlsx')
    print(f"[LOG] Successfully created the excel file: {excel_name}.xlsx")

def createWorksheet(ws, col, val):
    ws.cell(row=1, column=col, value=val).fill = green_fill


def get_week_of_month(date_obj):
    first_day = date_obj.replace(day=1)
    dom = date_obj.day

    adjusted_dom = dom + first_day.weekday()
    return int(math.ceil(adjusted_dom / 7.0))


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    exec()
