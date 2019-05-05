from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import xlsxwriter
from openpyxl import load_workbook
from openpyxl import Workbook
import os.path

def simple_get(url):
# Get content by user input url by making an HTTP GET request.
# Content type can be HTML/XML, check to see which one

    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during re quests to {0} : {1}'.formart(url, str(e)))
        return None

def is_good_response(resp):
# Determines whether response if HTML

    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200
            and content_type is not None
            and content_type.find('html') > -1)

def log_error(e):
    print(e)

def get_listings(city, raw_input):
    # Download the searched page of listed cars
    url_pre = 'https://'
    url_end = '.craigslist.org/search/cta?query='
    has_image = '&sort=rel&hasPic=1'

    # Checks for spaces and replaces with +
    if " " in raw_input:
        searched_car = raw_input.replace(" ", "+")
    else:
        searched_car = raw_input

    # Create standard url for searching specific vehicle
    combined_url = url_pre + city + url_end + searched_car + has_image

    # Obtains HTML response and checks if a response works
    response = simple_get(combined_url)
    if response is not None:
        html = BeautifulSoup(response, 'html.parser')
        listings = html.find_all('li', class_= 'result-row')
        return listings

def get_info(listings):
    # Lists of all needed info
    title = []
    pricing = []
    date = []
    id = []
    link = []

    for i in range(len(listings)):
        title.append(listings[i].find('a', class_= 'result-title hdrlnk').text)
        pricing.append(int(listings[i].find('span', class_= 'result-price').text.replace('$','')))
        date.append(listings[i].find('time', class_= 'result-date')['datetime'])
        id.append(int(listings[i].find('a', class_= 'result-title hdrlnk')['data-id']))
        link.append(listings[i].find('a', class_= 'result-title hdrlnk')['href'])

    return(title, pricing, date, id, link)


def filtered_search(raw_input, title, pricing, date, id, link):
    # Splits keywords into individual words
    car_model = raw_input.split()
    current_len = len(title)

    i = 0
    for k in car_model:
        while i < current_len:
            if k.upper() in title[i].upper():
                i += 1
            else:
                # Removes listings that does not contain searched keywords
                del title[i]
                del pricing[i]
                del date[i]
                del id[i]
                del link[i]

                current_len = len(title)
        i = 0 # Reset counter

        return(title, pricing, date, id, link)

def createNewWorksheet(title, pricing, date, id, link):
    # Creating a workbook and worksheet if no previous existed
    workbook = xlsxwriter.Workbook('ScrappedListings.xlsx')
    worksheet = workbook.add_worksheet('Listings')

    # Adding bold, money formats
    bold = workbook.add_format({'bold': True})
    money = workbook.add_format({'num_format': '$#,##0'})

    # Data Headers
    worksheet.write('A1', 'Date', bold)
    worksheet.write('B1', 'Listed Price', bold)
    worksheet.write('C1', 'ID', bold)
    worksheet.write('D1', 'Title', bold)
    worksheet.write('E1', 'Link', bold)

    # Starting from first cell below headers
    row = 1
    col = 0

    # Writes all scrapped info into specific columns
    for i in range(len(title)):
        worksheet.write(row, col, date[i])
        worksheet.write(row, col + 1, pricing[i], money)
        worksheet.write(row, col + 2, id[i])
        worksheet.write(row, col + 3, title[i])
        worksheet.write(row, col + 4, link[i])
        row += 1

    workbook.close()

def addListings(working_worksheet, title, pricing, date, id, link):
    # Determines how many rows are written in the excel sheet and
    # adds the amount of new listings to the list
    nRows = working_worksheet.max_row
    nAdd = len(title)

    i = 0
    for r in range(nRows, nRows + nAdd):
        working_worksheet.cell(row = r, column = 1).value = date[i]
        working_worksheet.cell(row = r, column = 2).value  = pricing[i]
        working_worksheet.cell(row = r, column = 3).value  = id[i]
        working_worksheet.cell(row = r, column = 4).value  = title[i]
        working_worksheet.cell(row = r, column = 5).hyperlink  = link[i]

        working_worksheet.cell(row = r, column = 2).number_format = u'"$ "#,##'
        working_worksheet.cell(row = r, column = 5).style  = 'Hyperlink'

        i += 1

    loaded_workbook.save('ScrappedListings.xlsx')
    loaded_workbook.close()

if __name__ == "__main__":
    major_cities = ['atlanta', 'houston', 'losangeles', 'sfbay', 'chicago', 'newyork',
        'seattle', 'orangecounty', 'sandiego', 'washingtondc', 'portland', 'boston',
        'phoenix', 'denver']
    raw_input = input('What vehicle? ')

    for current_city in major_cities:
        listings = get_listings(current_city, raw_input)
        (title, pricing, date, id, link) = get_info(listings)

        (up_title, up_pricing, up_date, up_id, up_link) =filtered_search(raw_input,
                                                                        title, pricing, date, id, link)

        exists = os.path.isfile('ScrappedListings.xlsx')
        if exists:
            loaded_workbook = load_workbook(filename = 'ScrappedListings.xlsx')
            loaded_worksheet = loaded_workbook.active
            addListings(loaded_worksheet, up_title, up_pricing, up_date, up_id, up_link)
        else:
            createNewWorksheet(up_title, up_pricing, up_date, up_id, up_link)
