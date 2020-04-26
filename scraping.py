from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as BS
import csv
import xlsxwriter
from urllib.parse import quote
import openpyxl
import time


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}
links = []
pages = [
    "https://ae-box.com/%D9%85%D9%83%D8%A7%D8%A6%D9%86-%D8%A7%D9%84%D9%82%D9%87%D9%88%D8%A9/c447415102",
    "https://ae-box.com/%D9%85%D9%83%D8%A7%D8%A6%D9%86-%D8%A7%D9%84%D9%82%D9%87%D9%88%D8%A9/c447415102?page=2",
    "https://ae-box.com/%D8%A3%D8%AC%D9%87%D8%B2%D8%A9-%D8%A7%D9%84%D9%85%D8%B7%D8%A8%D8%AE-%D8%A7%D9%84%D9%83%D9%87%D8%B1%D8%A8%D8%A7%D8%A6%D9%8A%D8%A9/c251785093",
    "https://ae-box.com/%D9%83%D8%A8%D8%B3%D9%88%D9%84%D8%A7%D8%AA-%D9%82%D9%87%D9%88%D8%A9/c2060306918",
    "https://ae-box.com/%D8%A3%D8%AC%D9%87%D8%B2%D8%A9-%D8%A5%D8%B2%D8%A7%D9%84%D8%A9-%D8%A7%D9%84%D8%B4%D8%B9%D8%B1-ipl/c1145434894",
    "https://ae-box.com/%D8%A7%D9%84%D8%B9%D9%86%D8%A7%D9%8A%D8%A9-%D8%A8%D8%A7%D9%84%D8%A7%D8%B3%D9%86%D8%A7%D9%86/c169123114",
    "https://ae-box.com/%D8%A7%D9%84%D8%B9%D9%86%D8%A7%D9%8A%D8%A9-%D8%A8%D8%A7%D9%84%D8%A8%D8%B4%D8%B1%D8%A9/c1484024240",
    "https://ae-box.com/%D8%A7%D9%84%D8%B9%D9%86%D8%A7%D9%8A%D8%A9-%D8%A8%D8%A7%D9%84%D8%B4%D8%B9%D8%B1/c710052017",
    "https://ae-box.com/%D8%A7%D9%84%D8%B9%D9%86%D8%A7%D9%8A%D8%A9-%D8%A8%D8%A7%D9%84%D8%B4%D8%B9%D8%B1/c710052017?page=2",
    "https://ae-box.com/%D9%85%D9%86%D8%AA%D8%AC%D8%A7%D8%AA-%D8%B7%D8%A8%D9%8A%D8%A9/c1525046448",
    "https://ae-box.com/%D9%85%D9%86%D8%AA%D8%AC%D8%A7%D8%AA-%D8%A7%D9%84%D8%B9%D9%86%D8%A7%D9%8A%D8%A9-%D8%A8%D8%A7%D9%84%D8%B1%D8%AC%D9%84/c1866982981",
    "https://ae-box.com/%D8%A3%D8%B7%D9%82%D9%85-%D9%84%D8%AD%D8%A7%D9%81%D8%A7%D8%AA/c1353324353",
    "https://ae-box.com/%D8%A8%D8%B7%D8%A7%D9%86%D9%8A%D8%A7%D8%AA/c578238018",
    "https://ae-box.com/%D8%B4%D9%86%D8%B7-%D9%86%D8%B3%D8%A7%D8%A6%D9%8A%D8%A9/c1312388172",
    "https://ae-box.com/%D9%85%D8%AD%D8%A7%D9%81%D8%B8-%D9%86%D8%B3%D8%A7%D8%A6%D9%8A%D8%A9/c403149645",
    "https://ae-box.com/%D8%A3%D8%B3%D8%A7%D9%88%D8%B1-%D9%88%D8%AA%D8%B9%D9%84%D9%8A%D9%82%D8%A7%D8%AA/c1978446414",
    "https://ae-box.com/%D8%A5%D9%83%D8%B3%D8%B3%D9%88%D8%A7%D8%B1%D8%A7%D8%AA-%D8%A7%D9%84%D8%B3%D9%8A%D8%A7%D8%B1%D8%A7%D8%AA/c1100248793",
    "https://ae-box.com/%D8%A7%D9%84%D8%B9%D9%86%D8%A7%D9%8A%D8%A9-%D8%A8%D8%A7%D9%84%D8%AD%D8%AF%D8%A7%D8%A6%D9%82/c1418414594",
    "https://ae-box.com/%D8%A3%D9%86%D8%AA%D8%B1%D9%86%D8%AA-%D9%88%D8%B4%D8%A8%D9%83%D8%A7%D8%AA/c637432153",
    "https://ae-box.com/%D8%A7%D9%84%D9%82%D8%B6%D8%A7%D8%A1-%D8%B9%D9%84%D9%89-%D8%A7%D9%84%D8%AD%D8%B4%D8%B1%D8%A7%D8%AA/c760064965",
    "https://ae-box.com/%D9%85%D8%B3%D8%AA%D9%84%D8%B2%D9%85%D8%A7%D8%AA-%D8%A7%D9%84%D8%B7%D9%81%D9%84-%D9%88%D8%A7%D9%84%D9%85%D8%B1%D8%A3%D8%A9/c403831448",
    "https://ae-box.com/%D8%A7%D9%84%D8%B5%D9%88%D8%AA%D9%8A%D8%A7%D8%AA-%D9%88%D8%A7%D9%84%D9%85%D8%B1%D8%A6%D9%8A%D8%A7%D8%AA/c1516782728",
]

class Find_item():

    def scrap_item(link):
        req = Request(url=(link), headers=headers)
        resp = urlopen(req).read()
        html = BS(resp, 'html.parser')
        html = html.find_all('div',class_ = 'product')
        return html


    def result_item(html_result):
        for i in html_result:
            links.append(i.a.get('href'))

    def add_link(pages):
        for page in pages:
            link = page
            html_result = Find_item.scrap_item(link)
            Find_item.result_item(html_result)



class Item():

    def main_scrap(link):
        req = Request(url=(link), headers=headers)
        resp = urlopen(req).read()
        html = BS(resp, 'html.parser')
        return html


    def item_result(html):
        title = html.find('h1',class_ = 'product-details__title brand-title')
        title = title.text   
        try:
            if html.find('p',class_ = 'product-details__price').find('span',class_ = 'price-after').text is not None:
                price = html.find('p',class_ = 'product-details__price').find('span',class_ = 'price-after').text
        except:
            price = html.find('p',class_ = 'product-details__price').span.text

        # try:
        #     description_1 = html.find('p',class_ = 'ql-align-right')
        #     description_1 = description_1.text
        # except:
        #     pass
        # try:
        #     description_2 = html.find('ol',class_ = '')
        #     d = []
        #     d = description_2.find_all('li')
        #     description_2 = []
        #     for i in d:
        #         description_2.append(i.text)
        # except:
        #     try:   
        #         description_2 = html.find('div',class_ = 'col-md-7')
        #         description_2 = description_2.find('p',class_ = 'ql-align-right').text
        #     except:
        #         pass
        try:

            item_id = html.find('p',class_ = 'product-details__price').span
            item_id = item_id.get('id').split('_')[2]

        except:
            item_id = html.find('p',class_ = 'product-details__price').span
            item_id = item_id.get('id').split('_')[1]

        photo = html.find('div',class_ = 'thumb product-carousel owl-carousel owl-theme').a
        photo = photo.img.get('src')
        try:
            brand = html.find('h1',class_ = 'product-details__title brand-title').a.img.get('alt')
        except:
            brand = 'none'

        try:
            quantity = html.find('span',class_ = 'not-available-large').text
        except:
            quantity = 'none'

        return title , price , item_id , photo , brand , quantity
        

Find_item.add_link(pages)     # item links
links_print = []
titles = []
prices = []
item_ids = []
photos = []
brands = []
quantitys = []
for i in links:
    link = quote(i,safe='/:?=&')
    html = Item.main_scrap(link=link)      # main data      
    result = Item.item_result(html)
    titles.append(result[0])
    prices.append(result[1])
    item_ids.append(result[2])
    photos.append(result[3])
    brands.append(result[4])
    quantitys.append(result[5])
    links_print.append(i)




workbook = xlsxwriter.Workbook('analytics.xlsx')
worksheet = workbook.add_worksheet()

def printresult():
    for i in range(len(titles)):

        # if i == 0:
        #     a = "A1"
        #     b = "B1"
        #     c = "C1"
        #     d = "D1"
        #     e = "E1"
        #     f = "F1"
        #     g = "G1"
        #     h = "H1"
        #     s = "I1"
        #     worksheet.write(a, "المعرف")
        #     worksheet.write(b, "العنوان")
        #     worksheet.write(c,"الوصف")
        #     worksheet.write(d, "الرابط")
        #     worksheet.write(e, "الحالة")
        #     worksheet.write(f, "السعر")
        #     worksheet.write(g, "مدى التوفر")
        #     worksheet.write(h,"رابط الصورة")
        #     worksheet.write(s, "العلامة التجارية")
        # else:
        a = "A" + str(int(i+1))
        b = "B" + str(int(i+1))
        c = "C" + str(int(i+1))
        d = "D" + str(int(i+1))
        e = "E" + str(int(i+1))
        f = "F" + str(int(i+1))
        g = "G" + str(int(i+1))
        h = "H" + str(int(i+1))
        s = "I" + str(int(i+1))
        worksheet.write(a, item_ids[i])
        worksheet.write(b, titles[i])
        worksheet.write(c, titles[i])
        worksheet.write(d, links_print[i])
        worksheet.write(e, "جديد")
        worksheet.write(f, prices[i])
        if quantitys[int(i)] == 'none':
            worksheet.write(g, "in stock")
        else:
            worksheet.write(g, "out of stock ")

        worksheet.write(h, photos[i])

        if brands[i] == "none":
            worksheet.write(s, "")
        else:
            worksheet.write(s, brands[i])


printresult()
workbook.close()


while True:
    book = openpyxl.load_workbook('analytics.xlsx')
    sheet = book.active
    quantitys = []
    for i in links:
        link = quote(i,safe='/:?=&')
        html = Item.main_scrap(link=link)
        try:
            quantitys.append(html.find('span',class_ = 'not-available-large').text)
        except:
            quantitys.append('none')
    for q in range(len(quantitys)):
        q = q + 1 
        if quantitys[int(q-1)] == 'none':
            g = "G" + str(int(q))
            sheet[g] = "in stock"
        else:
            g = "G" + str(int(q))
            sheet[g] = "out of stock"
    book.save('analytics.xlsx')
    time.sleep(20000‬)

