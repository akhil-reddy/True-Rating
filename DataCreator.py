''' The data file of this program will be created in the location where this program is located. '''

import requests
import xlsxwriter
from bs4 import BeautifulSoup

name=[]
date=[]
heading=[]
review=[]

workbook = xlsxwriter.Workbook('Data.xlsx')
worksheet = workbook.add_worksheet()

print ("Please wait! This may take a while (approximately 10 minutes)...")
    
for m in range(0,2670,10):

    url = "< Insert movie's review page URL >"+str(m)
    r = requests.get(url)
    soup = BeautifulSoup(r.content,"html.parser")

    g_data1 = soup.find_all("div", {"id" : "tn15content"})   

    for item in g_data1:
        for a in range(10):
            if m<2645:
                s1 = item.contents[11].find_all("a")[2*a+1].text
                name.append(s1)
    for item2 in g_data1:
        for b in range(10):
            if m<2645:
                s2 = item2.contents[11].find_all("h2")[b].text
                heading.append(s2)
    for item3 in g_data1:
        for c in range(10):
            if m<2645:
                s3 = item3.contents[11].find_all("p")[c].text
                review.append(s3)

    row=0
    col=0 
    for i in name:
        worksheet.write(row,col,i)
        row +=1

    row=0
    col=1
    for j in heading:
        worksheet.write(row,col,j)
        row +=1

    row=0
    col=2
    for k in review:
        worksheet.write(row,col,k)
        row +=1

workbook.close()
print("This program is executed.")