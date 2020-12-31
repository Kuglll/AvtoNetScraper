import requests
from bs4 import BeautifulSoup
import xlwt
from xlwt import Workbook
from datetime import datetime


#oblika
#6010 - supermoto
#6002 - enduro
#6016 - cross

#znamke
# KTM
# Husqvarna

baseUrl = "https://www.avto.net"
# ccmMax=750 - filtrira ven vse 950 SM je in 1000+ ccm adventurje
userUrl = "https://www.avto.net/Ads/results.asp?znamka=KTM&oblika=6010&EQ7=1110100120&EQ9=100000000&KAT=1060000000&ccmMax=750" #to be modified by user to choose Kategorija(avto, moto), Letnik, Znamka, oblika, ccmMax

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"}

page = requests.get(userUrl, headers=headers)

pageContent = BeautifulSoup(page.content, "html.parser")

#find vehicles
vehicles = pageContent.findAll("div", {"class": "row bg-white position-relative GO-Results-Row GO-Shadow-B"})
vehicleNames = pageContent.findAll("div", {"class": "GO-Results-Naziv bg-dark px-3 py-2 font-weight-bold text-truncate text-white text-decoration-none"})

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
    if ("DUKE" in currentModel.upper() or "LC4" in currentModel):  #TODO: modify to accept user input
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
            if(temp.find("th").get_text().strip() == "Letnik:"):
                year = temp.find("td").get_text().strip()
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
    id = year+kilometri+lastnik+avtohisa
    #if id in set - continue

    #Write to excel
    #numberOfRows = len(sheet._Worksheet__rows) //TODO: modify to only append to excel not overwrite (get length + 1)
    sheet.write(index + 1, 0, currentModel)
    sheet.write(index + 1, 1, year)
    sheet.write(index + 1, 2, kilometri)
    sheet.write(index + 1, 3, lastnik)
    sheet.write(index + 1, 4, avtohisa)
    sheet.write(index + 1, 5, price)
    sheet.write(index + 1, 6, vehicleUrl)
    sheet.write(index + 1, 7, datetime.today().strftime('%Y-%m-%d'))
    sheet.write(index + 1, 8, year+kilometri+lastnik+avtohisa)

    #user input
    #vsakemu vehiclu dodaj ID(letnik, kilometri, lastnik, avtohisa), ko poženeš skripto moraš preverit IDje
    #če je že v bazi, preveriš če se je cena kej spremenila
    #če ne ga appendaš na koncu baze

    print("*******************************NEXT RESULT******************************")

wb.save("test.xls")