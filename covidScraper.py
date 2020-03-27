import xlsxwriter
import xlrd
import requests
from bs4 import BeautifulSoup
from os.path import expanduser
import datetime

home = expanduser("~")
provData = {}
total = 0

URL = 'https://www.ctvnews.ca/health/coronavirus/tracking-every-case-of-covid-19-in-canada-1.4852102'
page = requests.get(URL)

soup = BeautifulSoup(page.content, 'html.parser')
provinces = soup.find_all('div', class_='covid-province-container')
for prov in provinces:
    tempName = prov.find('h2', class_='covid-head').text

    name = tempName[0:tempName.index("(")].strip()
    data = prov.find('td').text.strip()
    provData[name] = int(data)
    total = total + int(data)

repatriated = provinces[len(provinces) - 1]
tempData = repatriated.find_next_sibling("p").string
data = tempData[0:tempData.index(" ")].strip()

provData['Repatriated'] = data
provData['Total'] = total

# Open Data Worksheet and Store new Values
wbReader = xlrd.open_workbook("{}/Desktop/Test1.xlsx".format(home))
wbWriter = xlsxwriter.Workbook("{}/Desktop/Test1.xlsx".format(home))
oldSheet = wbReader.sheet_by_index(0)
newSheet = wbWriter.add_worksheet(oldSheet.name)
nrows = oldSheet.nrows
ncols = oldSheet.ncols

# Format Date Header in Row Zero
for col in range(ncols):
    newSheet.write(0, col, oldSheet.cell(0, col).value)
newSheet.write(0, ncols, datetime.datetime.now().strftime("%B %d, %I:%M %p"))

# Enter Data Points
for row in range(1, nrows - 2):
    provName = oldSheet.cell(row, 0).value
    if provName in provData:
        data = provData[provName]
    else:
        print("Not in dictionary")
        print(provName)
        data = -1
    for col in range(ncols):
        newSheet.write(row, col, oldSheet.cell(row, col).value)
    newSheet.write(row, ncols, data)

# Format Labels for New Case and Rate of New Case (static values)
newSheet.write(nrows - 2, 0, oldSheet.cell(nrows - 2, 0).value)
newSheet.write(nrows - 1, 0, oldSheet.cell(nrows - 1, 0).value)

# Enter Formulas For New Case and Rate of New Case Calculations
for col in range(2, ncols + 1):
    newCasesFormula = '=' + chr(65 + col) + str(nrows - 2) + '-' + chr(65 + col - 1) + str(nrows - 2)
    newSheet.write_formula(nrows - 2, col, newCasesFormula)

for col in range(3, ncols + 1):
    rateOfNewCasesFormula = '=' + chr(65 + col) + str(nrows - 1) + '-' + chr(65 + col - 1) + str(nrows - 1)
    newSheet.write_formula(nrows - 1, col, rateOfNewCasesFormula)

wbWriter.close()
