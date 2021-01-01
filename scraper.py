import requests
from bs4 import BeautifulSoup
import xlwt
from xlwt import Workbook
from datetime import datetime
from xlrd import open_workbook
from xlutils.copy import copy
import time


#oblika
#6010 - supermoto
#6002 - enduro
#6016 - cross

#znamke
# KTM
# Husqvarna

baseUrl = "https://www.avto.net"
userUrl = "https://www.avto.net/Ads/results.asp?EQ7=1110100120&EQ9=100000000"

print("Choose category (1-Avto, 2-Moto):")
category = input()
if(category == "2"):
    userUrl += "&KAT=1060000000"

print("Choose brand: (Audi, Volkswagen, KTM...). Must be same as on Avto.net.")
brand = input()
userUrl += "&znamka=" + brand

if(category == "1"):
    print("What model are you searching for?")
    model = input()
    userUrl += "&model=" + model

if(category == "2"):
    print("What are you searching for 6010 - supermoto, 6002 - enduro, 6016 - cross")
    motoType = input()
    userUrl += "&oblika=" + motoType

print("Choose min year (ex: 2012). Leave empty for all years.")
yearMin = input()
if(yearMin != ""):
    userUrl += "&letnikMin=" + yearMin


print("Choose max year (ex: 2018). Leave empty for all years.")
yearMax = input()
if(yearMax != ""):
    userUrl += "&letnikMax=" + yearMax

if(category == "1"):
    print("Choose min kW (ex: 77). Leave blank for any amount of min kW.")
    minKw = input()
    if(minKw != ""):
        userUrl += "&kwMin=" + minKw

    print("Choose max kW (ex: 200). Leave blank for any amount of max kW.")
    maxKw = input()
    if (maxKw != ""):
        userUrl += "&kwMax=" + maxKw

if(category == "2"):
    print("Choose min cmm: (ex: 125, leave blank for any min ccm)")
    minCcm = input()
    if(minCcm != ""):
        userUrl += "&ccmMin=" + minCcm

    print("Choose max ccm: (ex: 750, leave blank for any max ccm)")
    maxCcm = input()
    if(maxCcm != ""):
        userUrl += "&ccmMax=" + maxCcm

print("Choose min kilometers (ex: 100000, leave blank for any min kilometers):")
minKm = input()
if(minKm != ""):
    userUrl += "&kmMin=" + minKm

print("Choose max kilometers (ex: 250000, leave blank for any max kilometers):")
maxKm = input()
if (maxKm != ""):
    userUrl += "&kmMax=" + maxKm

print("Choose words to be filtered out of vehicle name. For example if you don't want any DUKE in your results, write DUKE.")
print("You can write multiple words, seperate them with comma (,). Ex: Duke,LC4,EXC")
wordsToBeFilteredOut = input()

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"}

page = requests.get(userUrl, headers=headers)

pageContent = BeautifulSoup(page.content, "html.parser")

#find vehicles
vehicles = pageContent.findAll("div", {"class": "row bg-white position-relative GO-Results-Row GO-Shadow-B"})
vehicleNames = pageContent.findAll("div", {"class": "GO-Results-Naziv bg-dark px-3 py-2 font-weight-bold text-truncate text-white text-decoration-none"})

set = set()

try:
    readOnlyWorkBook = open_workbook(brand+".xls")
    readSheet = readOnlyWorkBook.sheet_by_index(0)
    wb = copy(readOnlyWorkBook)
    sheet = wb.get_sheet(0)
    for i in range(1, len(sheet._Worksheet__rows)):
        set.add(readSheet.cell(i, 8).value)
except:
    print("File not found. Creating a new one.")
    wb = Workbook()
    sheet = wb.add_sheet("Sheet")
    sheet.write(0, 0, "Model")
    sheet.write(0, 1, "Letnik")
    sheet.write(0, 2, "Kilometri")
    sheet.write(0, 3, "Lastnik")
    sheet.write(0, 4, "Avto hiša")
    sheet.write(0, 5, "Cena")
    sheet.write(0, 6, "Link")
    sheet.write(0, 7, "Date added")
    sheet.write(0, 8, "ID")

for index, vehicle in enumerate(vehicles):
    #filter out models that you dont want
    currentModel = vehicleNames.__getitem__(index).get_text().strip()
    needToBreak = False
    if(wordsToBeFilteredOut != ""):
        for word in wordsToBeFilteredOut.split(","):
            if(word.upper() in currentModel.upper()):
                needToBreak = True

    if(needToBreak):
        continue

    #find vehicle url and parse each one to get more info
    vehicleUrl = baseUrl + BeautifulSoup(str(vehicle), "html.parser").find("a", {"class": "stretched-link"})["href"][2:]
    print("Url:" + vehicleUrl)
    vehiclePage = requests.get(vehicleUrl, headers=headers)
    vehiclePageContent = BeautifulSoup(vehiclePage.content, "html.parser")
    try:
        price = vehiclePageContent.find("p", {"class": "h2 font-weight-bold align-middle py-4 mb-0"}).get_text().strip()
    except:
        price = "/"
    podatki = vehiclePageContent.findAll("tr")
    year = "Novo"
    lastnik = "/"
    for podatek in podatki:
        try:
            temp = BeautifulSoup(str(podatek), "html.parser")
            #Year
            if(temp.find("th").get_text().strip() == "Letnik:" and category == "2"):
                year = temp.find("td").get_text().strip()
            if("Prva registracija" in temp.find("th").get_text().strip() and category == "1"):
                year = temp.find("td").get_text().strip().split("/")[0].strip()
        except:
            pass
        try:
            #Kilometri
            if (temp.find("th").get_text().strip() == "Prevoženi km:"):
                kilometri = temp.find("td").get_text().strip()
                if(kilometri == ""):
                    kilometri = "/"
        except:
            pass
        try:
            #Lastnik
            for extraPodatek in temp.findAll("li"):
                if("lastnik" in extraPodatek.get_text()):
                    lastnik = extraPodatek.get_text()
        except:
            pass

    #Avtohiša DA/NE
    try:
        vehiclePageContent.find("div", {"class": "col-12 text-center py-3"}).get_text().strip()
        avtohisa="DA"
    except:
        avtohisa="NE"

    #Get the ID of the vehicle
    id = year+kilometri+lastnik+avtohisa+currentModel
    if(id in set):
        for i in range(1, len(sheet._Worksheet__rows)):
            if(id == readSheet.cell(i, 8).value):
                tmp = readSheet.cell(i, 5).value.split("/")
                if(price != tmp[len(tmp)-1]):
                    sheet.write(i, 5, str(readSheet.cell(i, 5).value) + "/" + str(price))
                break
    else:
        #Write to excel
        numberOfRows = len(sheet._Worksheet__rows)
        sheet.write(numberOfRows, 0, currentModel)
        sheet.write(numberOfRows, 1, year)
        sheet.write(numberOfRows, 2, kilometri)
        sheet.write(numberOfRows, 3, lastnik)
        sheet.write(numberOfRows, 4, avtohisa)
        sheet.write(numberOfRows, 5, price)
        sheet.write(numberOfRows, 6, vehicleUrl)
        sheet.write(numberOfRows, 7, datetime.today().strftime('%Y-%m-%d'))
        sheet.write(numberOfRows, 8, year+kilometri+lastnik+avtohisa+currentModel)

    print("*******************************NEXT RESULT******************************")

try:
    wb.save(brand+".xls")
except:
    print("You need to close the excel file. The results are not saved. Try again. Closing in 3 seconds...")
    time.sleep(3)