#pip install beautifulsoup4
#pip install requests
#pip install win10toast
#pip install pandas
#pip install xlrd
#pip install unicode
#pip install Unidecode
#pip install encode
from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
from win10toast import ToastNotifier
from unidecode import unidecode
from datetime import date
import requests
import pandas as pd
import xlrd
import time
import os
import datetime
today = date.today() #šodien datums


d3 = today.strftime("%m/%d/%y") #datums ar mēnesi no sākuma pirmajās 2 pozīcijās
menesis = int(d3[0:2])

menesi = [
    'janvaris',
    'februaris',
    'marts',
    'aprilis',
    'maijs',
    'junijs',
    'julijs',
    'augusts',
    'septembris',
    'oktobris',
    'novembris',
    'decembris'
]

r = requests.get('http://jk.lv/dokumenti/nodarbibas/' + menesi[menesis-1] + '/KRN1.xls', auth=('jk', 'jk2000jk'))  #novelkam un saglabājam sarakstu
# open('~/saraksts.xls', 'wb').write(r.content)

#===============================================================================================================================================
filename = 'd:/saraksts.xls'
with open(filename, 'wb') as file:
    file.write(r.content)

wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_name('Sheet1')
tagad = datetime.datetime.now()

def intTryParse(value):
    try:
        int(value)
        return True
    except ValueError:
        return False
good_cells = list() #pārbaudam A kolonu, lai atrastu kur ir datumui un attiecīgi noteiktu kurā rindiņā ir 1. lekcija
good_indices = list() #nosakam rindiņas Nr kurā sākā aatiecīgās nedēļa lekcija
for i in range(sheet.nrows):
   if (intTryParse(sheet.cell_value(i,0))):
      good_cells.append(sheet.cell_value(i,0))
      good_indices.append(i)
      #if (sheet.cell_value(i,2)):
         #print(sheet.cell_value(i,2))

best_cell_ever = 0
for cell_index in good_indices:
    lekcijas_datums = xlrd.xldate_as_datetime(sheet.cell_value(cell_index,0),0).strftime("%d")
    if intTryParse(lekcijas_datums):
        if tagad.day <= int(lekcijas_datums): #meklējam nākamo lekciju datumu
            best_cell_ever = cell_index
            break
    else:
        print('Lekcijas datums nav int formata')
highest_number = 0
current_number = 1
i = best_cell_ever
lecture_list = list()

while True:
    if intTryParse(sheet.cell_value(i,1)): #meklējam int 2. kolonā, tiem pretī atiecīgās nedēļas lekcijas. Ja int lielāks par iepriekšējo, tad esam jau nonākuši pie nākamās nedēļas.
        current_number = int(sheet.cell_value(i,1))
        if current_number < highest_number:
            break
        else:
            highest_number = current_number
            lecture_list.append(sheet.cell_value(i,2)) #liekam listē attiecīgās dienas lekcija
    if i > best_cell_ever + 20: #ja nu kaut kādu iemeslu dēļu turpina skaitīt uz priekšu
        break
    i += 1

toaster = ToastNotifier()  #Windows 10 notification
toaster.show_toast("Lekcijas:", lecture_list[0] +" "+ lecture_list[1] +" "+lecture_list[2]+" "+lecture_list[3]+" "+lecture_list[4], icon_path=None, duration=5)