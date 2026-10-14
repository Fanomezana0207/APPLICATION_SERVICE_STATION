from pathlib import Path

import pandas

import openpyxl

import datetime

from service import Service




file = Path(r"C:\Users\FEHIZORO\Downloads\sales-summary-2026-10-07-2026-10-13.csv")



SUIVI_SERVICE_FILE = Path("D:/FANOMEZANA/SUIVI SERVICE/OCTOBRE/RECAP SERVICE_OCTOBRE_2026.xlsx")
FIRST_ROW = 4
LAVAGE_COL = 3
VIDANGE_COL = 6
VULCA_COL = 9
PARALLELISME_COL = 12

def selectRowByDate (date, sheet) :
    for i in range(4,sheet.max_row+1) :
        d = sheet.cell(row = i, column = 1).value
        d_en_str = d.strftime("%Y-%m-%d")
        print(f"date = {date}\t sheet.date = {d_en_str}\t=> {sheet.cell(row = i, column = 2).value}")
        if d_en_str == date:
            return i
    return f"date {date} non trouvé"  

def insertServiceToExcel(service, sheet, row) :
    if (service.getService() == "LAVAGE") :
        sheet.cell(row=row,column=LAVAGE_COL).value = service.getVente()
    elif (service.getService() == "VIDANGE") :
        sheet.cell(row=row,column=VIDANGE_COL).value = service.getVente()
    elif (service.getService() == "VULCA") :
        sheet.cell(row=row,column=VULCA_COL).value = service.getVente()
    elif (service.getService() == "PARALLELISME & CARROSSAGE") :
        sheet.cell(row=row,column=PARALLELISME_COL).value = service.getVente()

def insertSalesCategories(dataFilePath, outputFilePath) :
    dataFile = Path(dataFilePath.replace("\\","/"))
    data = pandas.read_csv(dataFile, encoding="utf-8-sig")
    data.columns = data.columns.str.strip() # Nettoyer les espaces éventuels
    outputFile = Path(outputFilePath.replace("\\","/"))

    if not outputFile.exists() :
        raise FileNotFoundError(outputFile)
        
    suivi_book = openpyxl.load_workbook(outputFile)
    print(suivi_book.sheetnames)
    #print("veuiller COPIER-COLLER le nom de la feuille choisi !")

    nom_feuille = input("veuiller COPIER-COLLER le nom de la feuille choisi : ")
    ws = suivi_book[nom_feuille]

    row = selectRowByDate(dataFile.stem[-10:],ws)

    for i in range(len(data)):
        serv = Service(data["Catégorie"][i],data["Ventes brutes"][i])
        insertServiceToExcel(serv, ws, row=row)

    suivi_book.save(outputFile)
    print("Modification enregisté")


dataFileToInput = input("Le lien du fichier sale-categorie à importer : ").replace("\"","")
dataFileOutput = input("Le lien du fichier Recap_service : ").replace("\"","")
insertSalesCategories(dataFileToInput, dataFileOutput)