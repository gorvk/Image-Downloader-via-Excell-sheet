import os
import xlrd
import xlwt
import requests
import urllib.request
from bs4 import BeautifulSoup
from xlwt import Workbook 


def downloadImgURL(url, i, naam, folderIndex):
   
    name=str(naam)+' '+str(i) 
    fullName=name+'.jpg' #name of file to save
    folderName="./images/"+str(folderIndex+1)+'.'+naam 
    folderPath=os.path.join(folderName, fullName) #path of folder
    urllib.request.urlretrieve(url, folderPath)
   
    #break

def spider():
    loc = ("./dataIn.xlsx") 
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    sheet.cell_value(0, 0)
    i=1
    #j=1
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet')
    rn=1
    cn=0
    ru=1
    cu=2
    sheet1.write(0, 0, 'Name of Items')
    sheet1.write(0, 2, 'URLs')     
    os.mkdir('images')
    for x in range(0, sheet.nrows-1):
        row=sheet.row_values(x)
        url="https://www.google.co.in/search?q="+str(row)+"&source=lnms&tbm=isch"
        sourceCode=requests.get(url)
        plainText=sourceCode.text
        
        
        parent_dir = "./images"
        # Path 
        path = os.path.join(parent_dir, str(x+1)+'.'+str(row[0])) 
          
        # Create the directory 
        # 'GeeksForGeeks' in 
        # '/home / User / Documents' 
        os.mkdir(path) 
        #print(folder)
        soup=BeautifulSoup(plainText,"html.parser")
        
        for link in soup.findAll('img', {'alt': 'Image result for '+str(row)}):
            src = link.get('src')
            downloadImgURL(src, i, row[0],x)         
            sheet1.write(rn, cn, row[0]+' '+str(i)) 
            i=i+1 
            sheet1.write(ru, cu, src) 
            rn+=1
            #cn+=1
            ru+=1
            #cu+=1
            wb.save('URLs.xls')

spider();    