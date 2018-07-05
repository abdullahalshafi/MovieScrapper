import requests
import xlwt
from bs4 import BeautifulSoup

details = "default"
downloadLink = "default"

# specify the url here
url = 'http://binodonmela.net/?product_cat=englishmovie&paged='
nextUrl = []
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

# specify the range or pages
for i in range(1, 50):
    finalUrl = url + str(i)
    source = requests.get(finalUrl)
    plain_text = source.text
    soup = BeautifulSoup(plain_text, 'html.parser')
    for link in soup.find_all('a', class_='button product_type_simple ajax_add_to_cart'):
        nextUrl.append(link.get('href'))

row = 0
for l in nextUrl:
    data = requests.get(l)
    plain_text1 = data.text
    soup1 = BeautifulSoup(plain_text1, 'html.parser')
    movieName = soup1.find('h1', class_='product_title entry-title').text
    for movieDownloadLink in soup1.find_all('div', class_='download'):
        downloadLink = movieDownloadLink.a.get('href')

    td = soup1.find('td', id='imdbimg')
    if td is not None:
        details = soup1.find("td", id='imdbimg').find_next_sibling("td").text

    sheet1.write(row, 0, movieName.strip())
    sheet1.write(row, 1, downloadLink)
    sheet1.write(row, 2, details)
    row = row + 1
    details = "none"

book.save("movies.xls")
print('done')
