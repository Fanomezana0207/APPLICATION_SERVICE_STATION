from app import *

print("Insérer le ventes par catégories, en une journée, dans le suivi de services")
dataFileToInput = input("Le lien du fichier sale-categorie à importer : ").strip().strip("\"")
dataFileOutput = input("Le lien du fichier Recap_service : ").strip().strip("\"")
insertSalesCategories(dataFileToInput, dataFileOutput)