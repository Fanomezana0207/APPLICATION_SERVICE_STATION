from app import *

print("Générer un fichier reçu Excel, LAVAGE")
dataFilePath = input("Le lien du fichier receipts à importer : ").strip().strip("\"")
generateReceiptsLavage(dataFilePath)