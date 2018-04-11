import requests
from bs4 import BeautifulSoup
import time
import xlsxwriter
from dateutil import parser
url = "https://www.homeshop18.com/account/orderDetails"
from variable import *
workbook = xlsxwriter.Workbook("HS18.xlsx")
workbook.formats[0].set_font_size(8)

worksheet = workbook.add_worksheet()
payload = ""
formate=workbook.add_format()
formate.set_bold()
formate.set_font_color('black')

worksheet.set_column(0, 0, 10)
worksheet.set_column(1,1,16)
worksheet.set_column(2,2,18)
worksheet.set_column(3,3,10)
worksheet.set_column(4,4,40)
worksheet.set_column(5,5,9)
worksheet.set_column(6,6,32)
worksheet.set_column(7,7,16)

worksheet.write_string  (0, 0,     "pid",formate)
worksheet.write_string  (0, 1,     "Date" ,formate)
worksheet.write_string  (0, 2,     "Name",formate)
worksheet.write_string  (0, 3,     "Contact",formate)
worksheet.write_string  (0, 4,     "Ship Detail",formate)
worksheet.write_string  (0, 5,     "Price",formate)
worksheet.write_string  (0, 6,     "Product Summary",formate)
worksheet.write_string  (0, 7,     "Company",formate)


row=1
for i in range(0,1000):
    orderId=orderId+2
    querystring = {"orderId":orderId}

    response = requests.request("GET", url, data=payload, headers=headers, params=querystring)
    soup = BeautifulSoup(response.content)
    notFound=soup.find_all('div',class_="sorry-content")
    if notFound:
        print(i)
    else:
        print(i)
        orderDetails = soup.find_all('div',class_="order-section-row1 clearfix")
        pid=orderDetails[0].find_all('div')[0].find('p').text
        pdate=orderDetails[0].find_all('div')[1].find('p').text
        dt = parser.parse(pdate)
        dt=(str(dt.year)+"-"+str(dt.month)+"-"+str(dt.day) )
        # print(dt)
        # time.sleep(10)
        shippingDetails= soup.find_all('div',class_="col col-shiping")
        # print(shippingDetails)
        shipdetail=' '.join( shippingDetails[0].find('p').text.split())
        shipdetail=shipdetail.split("Mobile")[0].title()
        # print(shipdetail)


        Mobile=''.join( shippingDetails[0].find('p').text.split()).split(':')[-1]

        name=shippingDetails[0].find('p').find('strong').text
        # print(name)

        paymentDetails=soup.find_all('div',class_="col col-payment last")
        paymentDetail=''.join( paymentDetails[0].find('p').text.split()).split(':')[-1]
        paymentDetail="Rs."+paymentDetail.split('.')[0]
        # paymentDetail="Rs."+str(paymentDetail)
        print(paymentDetail)
        productSummary=soup.find_all('td',valign="middle")
        productsumm=''
        for product in productSummary:
            if product.find('a'):
                productsumm=productsumm+product.find('a').text
        # print(pid)
        # print(productsumm)
        worksheet.write_string  (row, 0,     pid)
        worksheet.write_string  (row, 1,     pdate )
        worksheet.write_string  (row, 2,     name)
        worksheet.write_string  (row, 3,     Mobile)
        worksheet.write_string  (row, 4,     shipdetail)
        worksheet.write_string  (row, 5,     paymentDetail)
        worksheet.write_string  (row, 6,     productsumm)
        worksheet.write_string  (row, 7,     "HOMESHOP")
        
        
        row=row+1
workbook.close()     
