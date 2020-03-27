import requests
from bs4 import BeautifulSoup


def get_ncovlive_data():
    prov_data = {}

    url = 'https://ncov2019.live/data'
    page = requests.get(url)

    soup = BeautifulSoup(page.content, 'html.parser')
    provinces = soup.find("table", id="sortable_table_Canada").tbody.findAll("tr")

    # Remove and process the unique row for total cases at the top of the table
    total_cases_row = provinces.pop(0)
    data = total_cases_row.find("td", class_="text--green").text.strip()
    prov_data["Total"] = int(data.replace(',', ''))

    for prov in provinces:
        name = prov.find("td", class_="text--gray").contents[2].strip()
        data = prov.find("td", class_="text--green").text.strip()
        prov_data[name] = int(data.replace(',', ''))

    return prov_data
