from app import *

print("Insérer le ventes par catégories, en plusieurs jours, dans le suivi de services")
fileDir = input("Le lien du dossier où se trouve les fichiers : ").strip().strip("\"")
dataFileOutput = input("Le lien du fichier Recap_service : ").strip().strip("\"")
insertAllSalesCategories(fileDir, dataFileOutput)