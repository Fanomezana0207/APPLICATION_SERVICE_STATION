from app import *

print("Générer un fichier reçu Excel, facile à analyser")
dataFilePath = input("Le lien du fichier receipts à importer : ").strip().strip("\"")
generateReceipts(dataFilePath)