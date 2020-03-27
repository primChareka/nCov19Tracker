import requests
from bs4 import BeautifulSoup


def get_ctv_data():
    prov_data = {}
    total = 0

    url = 'https://www.ctvnews.ca/health/coronavirus/tracking-every-case-of-covid-19-in-canada-1.4852102'
    page = requests.get(url)

    soup = BeautifulSoup(page.content, 'html.parser')
    provinces = soup.find_all('div', class_='covid-province-container')
    for prov in provinces:
        temp_name = prov.find('h2', class_='covid-head').text
        name = temp_name[0:temp_name.index("(")].strip()
        data = prov.find('td').text.strip()
        prov_data[name] = int(data)
        total = total + int(data)

    repatriated = provinces[len(provinces) - 1]
    temp_data = repatriated.find_next_sibling("p").string
    data = temp_data[0:temp_data.index(" ")].strip()

    prov_data['Repatriated'] = int(data)
    prov_data['Total'] = total

    return prov_data