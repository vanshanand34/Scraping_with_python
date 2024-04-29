import openpyxl
from bs4 import BeautifulSoup
import requests
import re
import json

def getData(myurl):
    print(myurl)
    links = []
    images = []
    text = ""
    myhtml = requests.get(myurl).text
    parser = BeautifulSoup(myhtml,'html.parser')
    for link in parser.find_all('a',attrs={'href':re.compile('^https://')}):   #extracting all the href links
        links.append(link.get('href'))
    for img in parser.find_all('img',attrs={'src':re.compile('.png')}):         #extracting image sources
        images.append(img.get('src'))
    try:
        text = parser.find('body').text
        text = re.sub('[\n]+','\n',text)        #extracting text from body tag of html
    except AttributeError:
        text = ""
    return links,images,text


worksheet = openpyxl.load_workbook('Scrapping.xlsx')
obj = worksheet.active
res = []
mylinks = []
maxrow = obj.max_row
maxcol = obj.max_column
print(maxcol , maxrow)
for i in range(1,maxrow+1):
    for j in range(1,maxcol+1):
        cell = obj.cell(row=i,column=j)
        mylinks.append(str(cell.value))                 #extracting the urls from the given excel sheet
        

#extracting urls , images and text from each url given in the excell sheet and storing the data in a python dictionary
for i in mylinks:
    links,images,text = getData(i)
    res.append({"WebPageLink":i , "urls":links , "images":images , "text":text})
res = {"data":res}

#converting and storing data in the form of a json file
with open("myfile.json",'w') as f:
    json.dump(res,f)


