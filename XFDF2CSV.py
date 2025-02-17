# Importation des modules nécessaires
import xml.etree.ElementTree as ET
import csv
import os
import tkinter as tk
from tkinter import filedialog

# Ordre des colonnes tel que spécifié
columns_order = [
    "A-Name", "Department",
    "Q1-IT", "Q1-Comptabilite", "Q1-Multimedia", "Q1-Gestion de projet",
    "Q1-Communication", "Q1-Editorial", "Q1-Administration",
    "Q2-Name1", "Q2-Name2", "Q2-Name3", "Q2-Name4", "Q2-Name5",
    "Q2-Name6", "Q2-Name7", "Q2-Name8", "Q2-Name9",
    "Q3-IT", "Q3-Comptabilite", "Q3-Multimedia", "Q3-Gestion de projet",
    "Q3-Communication", "Q3-Editorial", "Q3-Administration",
    "Q4-IT", "Q4-Comptabilite", "Q4-Multimedia", "Q4-Gestion de projet",
    "Q4-Communication", "Q4-Editorial", "Q4-Administration"
]

# Fonction pour traiter les fichiers XFDF dans un répertoire et générer un CSV horizontal
def xfdf_folder_to_horizontal_csv(input_folder, output_csv_file, columns_order):
    try:
        # Créer une liste pour stocker toutes les lignes de données
        all_rows = []

        # Itérer sur tous les fichiers .xfdf dans le dossier
        for file_name in os.listdir(input_folder):
            if file_name.endswith(".xfdf"):
                file_path = os.path.join(input_folder, file_name)
                
                # Analyser chaque fichier XFDF
                tree = ET.parse(file_path)
                root = tree.getroot()

                # Créer un dictionnaire pour stocker les valeurs
                data_row = {col: "" for col in columns_order}

                # Extraire les champs et remplir le dictionnaire
                for field in root.findall(".//{http://ns.adobe.com/xfdf/}field"):
                    field_name = field.get("name")
                    field_value = field.find("{http://ns.adobe.com/xfdf/}value").text if field.find("{http://ns.adobe.com/xfdf/}value") is not None else ""
                    if field_name in data_row:
                        data_row[field_name] = field_value

                # Ajouter la ligne traitée à la liste
                all_rows.append(data_row)

        # Écrire toutes les données dans un fichier CSV avec un point-virgule comme délimiteur
        with open(output_csv_file, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=columns_order, delimiter=';')
            writer.writeheader()
            writer.writerows(all_rows)

        print(f"Fichier CSV horizontal consolidé généré avec succès : {output_csv_file}")
    except Exception as e:
        print(f"Erreur lors du traitement des fichiers : {e}")

# Fonction principale pour exécuter le programme
def main():
    # Créer l'objet racine Tkinter et le cacher
    root = tk.Tk()
    root.withdraw()

    # Demander à l'utilisateur de sélectionner le dossier d'entrée
    input_folder = filedialog.askdirectory(title="Sélectionnez le dossier contenant les fichiers XFDF")
    if not input_folder:
        print("Aucun dossier d'entrée sélectionné.")
        return

    # Demander à l'utilisateur de sélectionner le fichier CSV de sortie
    output_csv_file = filedialog.asksaveasfilename(
        title="Sélectionnez l'emplacement du fichier CSV de sortie",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
    )
    if not output_csv_file:
        print("Aucun fichier de sortie sélectionné.")
        return

    # Convertir les fichiers XFDF du dossier en CSV avec un format horizontal
    xfdf_folder_to_horizontal_csv(input_folder, output_csv_file, columns_order)

if __name__ == "__main__":
    main()