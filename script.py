from pathlib import Path

import pandas

from service import Service

if __name__ == "__main__":    

    CURENT_DIR = Path("__FILE__").parent.resolve()

    FILES_DIR = CURENT_DIR / "FILES" / "VENTE_PAR_CATEGORIE"

    file_1 = "FILES/VENTE_PAR_CATEGORIE/category-sales-summary-2026-10-01-2026-10-01.xlsx"

    for file in FILES_DIR.iterdir() :
        print(file)
        print(file.name)

    #data = pandas.read_excel(file_1)
    data = pandas.read_excel(file_1, sheet_name="category-sales-summary-2026-10-", usecols="A, C")

    #print(data)

    list_service = []
    for i in range(len(data)):
        if i != 4 :
            serv = Service(data["CatÃ©gorie"][i],data["Ventes brutes"][1])
            list_service.append(serv)

    for service in list_service:
        print(service)
