import xlsxwriter
import xlrd
from os.path import expanduser
import datetime
from scraper_ctv import get_ctv_data
from scraper_nCovLive import get_ncovlive_data

HOME = expanduser("~")


def store_data_in_excel(prov_data, wb_reader, wb_writer, sheet_name):

    old_sheet = wb_reader.sheet_by_name(sheet_name)
    new_sheet = wb_writer.add_worksheet(old_sheet.name)
    nrows = old_sheet.nrows
    ncols = old_sheet.ncols

    # Format Date Header in Row Zero
    for col in range(ncols):
        new_sheet.write(0, col, old_sheet.cell(0, col).value)
    new_sheet.write(0, ncols, datetime.datetime.now().strftime("%B %d, %I:%M %p"))

    # Enter Data Points
    for row in range(1, nrows - 2):
        prov_name = old_sheet.cell(row, 0).value
        if prov_name in prov_data:
            data = prov_data[prov_name]
        else:
            print("Not in dictionary")
            print(prov_name)
            data = -1
        for col in range(ncols):
            new_sheet.write(row, col, old_sheet.cell(row, col).value)
        new_sheet.write(row, ncols, data)

    # Format Labels for New Case and Rate of New Case (static values)
    new_sheet.write(nrows - 2, 0, old_sheet.cell(nrows - 2, 0).value)
    new_sheet.write(nrows - 1, 0, old_sheet.cell(nrows - 1, 0).value)

    # Enter Formulas For New Case and Rate of New Case Calculations
    for col in range(2, ncols + 1):
        newCasesFormula = '=' + chr(65 + col) + str(nrows - 2) + '-' + chr(65 + col - 1) + str(nrows - 2)
        new_sheet.write_formula(nrows - 2, col, newCasesFormula)

    for col in range(3, ncols + 1):
        rateOfNewCasesFormula = '=' + chr(65 + col) + str(nrows - 1) + '-' + chr(65 + col - 1) + str(nrows - 1)
        new_sheet.write_formula(nrows - 1, col, rateOfNewCasesFormula)


# Open Data Worksheet and Store new Values
reader = xlrd.open_workbook("{}/Desktop/CovidData.xlsx".format(HOME))
writer = xlsxwriter.Workbook("{}/Desktop/CovidData.xlsx".format(HOME))

province_data = get_ctv_data()
store_data_in_excel(province_data, reader, writer, "ctv-data")

province_data = get_ncovlive_data()
store_data_in_excel(province_data, reader, writer, "nCovLive-data")
writer.close()
