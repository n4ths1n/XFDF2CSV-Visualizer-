# Importation des modules nécessaires
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import xml.etree.ElementTree as ET
import csv
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.patches import Patch
import seaborn as sns
import networkx as nx

# Fonction utilitaire pour limiter à 20 catégories : retourne les 19 premières et regroupe le reste sous "Other"
def limit_top_20(series):
    series = series.sort_values(ascending=False)
    if len(series) <= 20:
        return series
    top19 = series.iloc[:19]
    other_sum = series.iloc[19:].sum()
    return pd.concat([top19, pd.Series([other_sum], index=["Other"])])

# Classe principale gérant l'application de visualisation et conversion CSV
class VisualisateurCSV:
    # Constantes pour le zoom et la taille des nœuds dans le graphique réseau
    BASE_TAILLE_NOEUD = 30
    ZOOM_IN_FACTOR = 1.1
    ZOOM_OUT_FACTOR = 0.9
    MIN_ZOOM = 0.5
    MAX_ZOOM = 5.0

    # Constructeur de la classe : initialisation de l'interface et des variables de l'application
    def __init__(self, racine):
        self.racine = racine
        self.racine.title("XFDF2CSV Visualizer")
        self.df = None                # DataFrame contenant les données CSV chargées
        self.G = None                 # Graphe utilisé pour la visualisation réseau
        self.pos = None               # Positions des nœuds dans le graphe
        self.liste_a_names = []       # Liste des noms utilisés pour la coloration des nœuds
        self.type_visu_actuelle = "barres"  # Type de visualisation par défaut
        self.echelle_actuelle = 1.0         # Facteur d'échelle initial pour le zoom
        # Dictionnaire contenant les textes des questions à afficher
        self.questions = {
            "Q1": "AVEC QUEL DÉPARTEMENT AIMERAIS-TU TRAVAILLER?",
            "Q2": "AVEC QUI AIMERAIS-TU TRAVAILLER EN DEHORS DE TON DÉPARTEMENT?",
            "Q3": "AVEC QUEL DÉPARTEMENT AIMERAIS-TU PAS TRAVAILLER?",
            "Q4": "À TON AVIS, QUEL DÉPARTEMENT NE TROUVERAIT PAS D'INTÉRÊT PROFESSIONNEL À TRAVAILLER AVEC TOI?",
            "Department": "DÉPARTEMENTS QUI ONT LE PLUS RÉPONDU"
        }
        # Configuration de l'interface graphique
        self.configurer_interface()

    # Méthode pour configurer l'interface graphique de l'application
    def configurer_interface(self):
        # Création du panneau gauche pour les boutons et options
        panneau_gauche = tk.Frame(self.racine, width=200, bg="#f0f0f0")
        panneau_gauche.pack(side=tk.LEFT, fill=tk.Y)
        # Création du panneau droit pour l'affichage des graphiques
        panneau_droit = tk.Frame(self.racine)
        panneau_droit.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH)
        # Label affichant la question courante
        self.lbl_question = tk.Label(panneau_droit, text="", font=('Arial', 10), wraplength=600)
        self.lbl_question.pack(pady=10)
        # Cadre pour la légende (utilisé pour le graphique réseau)
        self.frame_legende = tk.Frame(panneau_droit)
        self.lbl_legende = tk.Label(self.frame_legende, text="", font=('Arial', 9))
        self.lbl_legende.pack()
        self.frame_legende.pack(pady=5)
        # Bouton pour convertir les fichiers XFDF en CSV
        btn_convertir = tk.Button(panneau_gauche, text="Convertir XFDF en CSV", command=self.convert_xfdf_to_csv)
        btn_convertir.pack(pady=10, padx=10, fill=tk.X)
        # Bouton pour charger un fichier CSV
        btn_charger = tk.Button(panneau_gauche, text="Charger CSV", command=self.charger_csv)
        btn_charger.pack(pady=10, padx=10, fill=tk.X)
        # Création d'une combobox pour sélectionner la question à visualiser
        self.var_question = tk.StringVar()
        lbl_selection = tk.Label(panneau_gauche, text="Visualiser:", bg="#f0f0f0")
        lbl_selection.pack(pady=(20,5))
        self.combo_questions = ttk.Combobox(
            panneau_gauche,
            textvariable=self.var_question,
            values=["Department", "Q1", "Q2", "Q3", "Q4"],
            state="readonly"
        )
        self.combo_questions.pack(padx=10, fill=tk.X)
        self.combo_questions.current(0)
        self.combo_questions.bind("<<ComboboxSelected>>", self.actualiser_affichage)
        # Boutons pour sélectionner le type de visualisation
        self.btn_barres = tk.Button(panneau_gauche, text="Graphique en Barres", command=lambda: self.changer_visu("barres"))
        self.btn_barres.pack(pady=5, padx=10, fill=tk.X)
        self.btn_heatmap = tk.Button(panneau_gauche, text="Matrice de Chaleur", command=lambda: self.changer_visu("heatmap"))
        self.btn_heatmap.pack(pady=5, padx=10, fill=tk.X)
        self.btn_reseau = tk.Button(panneau_gauche, text="Réseau de Relations", command=lambda: self.changer_visu("reseau"))
        self.btn_reseau.pack(pady=5, padx=10, fill=tk.X)
        self.btn_pie = tk.Button(panneau_gauche, text="Graphique en Secteurs", command=lambda: self.changer_visu("pie"))
        self.btn_pie.pack(pady=5, padx=10, fill=tk.X)
        self.btn_line = tk.Button(panneau_gauche, text="Graphique en Lignes", command=lambda: self.changer_visu("line"))
        self.btn_line.pack(pady=5, padx=10, fill=tk.X)
        # Création de la figure Matplotlib pour l'affichage des graphiques
        self.figure = plt.Figure(figsize=(10,7), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=panneau_droit)
        self.canvas.get_tk_widget().pack(expand=True, fill=tk.BOTH)
        # Ajout de la barre d'outils de navigation Matplotlib
        self.toolbar = NavigationToolbar2Tk(self.canvas, panneau_droit)
        self.toolbar.update()
        # Connexion de l'événement de défilement pour gérer le zoom
        self.canvas.mpl_connect("scroll_event", self.gestion_zoom)

    # Méthode pour convertir les fichiers XFDF en CSV
    def convert_xfdf_to_csv(self):
        # Demande à l'utilisateur de sélectionner le dossier contenant les fichiers XFDF
        input_folder = filedialog.askdirectory(title="Sélectionnez le dossier contenant les fichiers XFDF")
        if not input_folder:
            messagebox.showinfo("Information", "Aucun dossier d'entrée sélectionné.")
            return
        # Demande à l'utilisateur de sélectionner l'emplacement et le nom du fichier CSV de sortie
        output_csv_file = filedialog.asksaveasfilename(
            title="Sélectionnez l'emplacement du fichier CSV de sortie",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not output_csv_file:
            messagebox.showinfo("Information", "Aucun fichier de sortie sélectionné.")
            return
        # Conversion des fichiers XFDF en CSV
        self.xfdf_folder_to_horizontal_csv(input_folder, output_csv_file)

    # Méthode pour traiter un dossier de fichiers XFDF et générer un CSV horizontal
    def xfdf_folder_to_horizontal_csv(self, input_folder, output_csv_file):
        # Ordre des colonnes dans le fichier CSV
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
        try:
            all_rows = []
            # Parcours de tous les fichiers du dossier
            for file_name in os.listdir(input_folder):
                if file_name.lower().endswith(".xfdf"):
                    file_path = os.path.join(input_folder, file_name)
                    tree = ET.parse(file_path)
                    root = tree.getroot()
                    # Initialisation d'un dictionnaire pour chaque ligne du CSV
                    data_row = {col: "" for col in columns_order}
                    # Extraction des valeurs des champs XFDF
                    for field in root.findall(".//{http://ns.adobe.com/xfdf/}field"):
                        field_name = field.get("name")
                        value_element = field.find("{http://ns.adobe.com/xfdf/}value")
                        field_value = value_element.text if value_element is not None else ""
                        if field_name in data_row:
                            data_row[field_name] = field_value
                    all_rows.append(data_row)
            # Écriture des données dans le fichier CSV
            with open(output_csv_file, mode='w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=columns_order, delimiter=';')
                writer.writeheader()
                writer.writerows(all_rows)
            messagebox.showinfo("Conversion réussie", f"Fichier CSV généré avec succès : {output_csv_file}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement des fichiers : {e}")

    # Méthode pour charger un fichier CSV et mettre à jour la visualisation
    def charger_csv(self):
        fichier = filedialog.askopenfilename(filetypes=[("Fichiers CSV", "*.csv")])
        if fichier:
            try:
                self.df = pd.read_csv(fichier, sep=";")
                # Mise à jour de la liste des noms pour la coloration dans le graphique réseau
                if "A-Name" in self.df.columns:
                    self.liste_a_names = self.df["A-Name"].unique().tolist()
                else:
                    self.liste_a_names = []
                self.actualiser_affichage()
            except Exception as e:
                messagebox.showerror("Erreur de chargement", f"Impossible de charger le fichier CSV.\n{str(e)}")

    # Méthode pour préparer les données en fonction de la question sélectionnée
    def preparer_donnees(self):
        if self.df is None:
            return None
        question = self.var_question.get()
        if question == "Department":
            # Comptage des réponses par département
            df_dept = self.df["Department"].value_counts().reset_index()
            df_dept.columns = ["Department", "count"]
            return df_dept
        elif question == "Q2":
            # Transformation des colonnes Q2 en un format long
            colonnes = [f"Q2-Name{i}" for i in range(1, 10) if f"Q2-Name{i}" in self.df.columns]
            if not colonnes:
                return None
            df_fusion = self.df.melt(id_vars=["A-Name", "Department"], value_vars=colonnes, value_name="Réponse")
            df_fusion = df_fusion[df_fusion["Réponse"] != "----"]
            return df_fusion
        else:
            # Transformation des colonnes Q1, Q3, Q4 en un format long et filtrage sur "Oui"
            prefixe = f"{question}-"
            colonnes = [col for col in self.df.columns if col.startswith(prefixe)]
            if not colonnes:
                return None
            df_fusion = self.df.melt(id_vars=["A-Name", "Department"], value_vars=colonnes, var_name="Catégorie", value_name="Réponse")
            df_fusion = df_fusion[df_fusion["Réponse"] == "Oui"]
            return df_fusion

    # Méthode pour afficher la visualisation selon le type sélectionné
    def afficher_visualisation(self, type_visu):
        if self.df is None:
            return
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)
        donnees = self.preparer_donnees()
        if donnees is None or donnees.empty:
            self.ax.text(0.5, 0.5, "Aucune donnée à afficher", ha="center", va="center")
            self.canvas.draw()
            return
        try:
            # Sélection du type de graphique à afficher
            if type_visu == "barres":
                self._afficher_barres(donnees)
            elif type_visu == "heatmap":
                self._afficher_heatmap(donnees)
            elif type_visu == "reseau":
                self._afficher_reseau(donnees)
            elif type_visu == "pie":
                self._afficher_pie(donnees)
            elif type_visu == "line":
                self._afficher_line(donnees)
            self.figure.tight_layout()
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Erreur de visualisation", f"Erreur de visualisation : {str(e)}")

    # Méthode privée pour afficher un graphique en barres
    def _afficher_barres(self, donnees):
        question = self.var_question.get()
        if question == "Department":
            donnees = donnees.sort_values("count", ascending=False)
            if len(donnees) > 20:
                top19 = donnees.iloc[:19]
                other_count = donnees["count"].iloc[19:].sum()
                donnees = pd.concat([top19, pd.DataFrame([["Other", other_count]], columns=["Department", "count"])], ignore_index=True)
            sns.barplot(x='count', y='Department', data=donnees, ax=self.ax, palette="viridis")
            self.ax.set_xlabel("Nombre de répondants")
        elif question == "Q2":
            comptage = donnees["Réponse"].value_counts().sort_values(ascending=False)
            if len(comptage) > 20:
                top19 = comptage.iloc[:19]
                other_count = comptage.iloc[19:].sum()
                comptage = pd.concat([top19, pd.Series([other_count], index=["Other"])])
            comptage.plot(kind="bar", ax=self.ax, color="skyblue")
            self.ax.set_ylabel("Mentions")
        else:
            comptage = donnees["Catégorie"].value_counts().sort_values(ascending=False)
            if len(comptage) > 20:
                top19 = comptage.iloc[:19]
                other_count = comptage.iloc[19:].sum()
                comptage = pd.concat([top19, pd.Series([other_count], index=["Other"])])
            sns.barplot(x=comptage.values, y=comptage.index, ax=self.ax, palette="rocket")
            self.ax.set_ylabel("Réponses 'Oui'")
        self.ax.tick_params(axis="x", rotation=45)

    # Méthode privée pour afficher une matrice de chaleur (pour Q2)
    def _afficher_heatmap(self, donnees):
        question = self.var_question.get()
        if question == "Q2":
            matrice = pd.crosstab(donnees["A-Name"], donnees["Réponse"])
            sns.heatmap(matrice, ax=self.ax, cmap="YlGnBu", cbar_kws={'label': 'Mentions'})
            self.ax.tick_params(axis="x", rotation=45)

    # Méthode privée pour afficher un graphique réseau
    def _afficher_reseau(self, donnees):
        self.G = nx.from_pandas_edgelist(donnees, "A-Name", "Réponse")
        self.pos = nx.spring_layout(self.G, k=0.3)
        self._redessiner_reseau()

    # Méthode privée pour redessiner le graphique réseau (utile lors du zoom)
    def _redessiner_reseau(self):
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)
        taille_noeud = self.BASE_TAILLE_NOEUD * self.echelle_actuelle
        # Définition des couleurs pour chaque nœud
        couleurs = []
        for node in self.G.nodes():
            if node in self.liste_a_names:
                couleurs.append("#1f77b4")
            else:
                couleurs.append("#2ca02c")
        nx.draw(self.G, self.pos, ax=self.ax, with_labels=True, node_size=taille_noeud, font_size=6, node_color=couleurs, edge_color="gray", width=0.5)
        # Création de la légende du graphique réseau
        legend_elements = [Patch(facecolor='#1f77b4', edgecolor='black', label='A-Name'),
                           Patch(facecolor='#2ca02c', edgecolor='black', label='Autres noms')]
        self.ax.legend(handles=legend_elements, loc='upper right')
        self.canvas.draw()
        self.lbl_legende.config(text="Bleu - A-Name   |   Vert - Autres participants")

    # Méthode privée pour afficher un graphique en secteurs
    def _afficher_pie(self, donnees):
        question = self.var_question.get()
        if question == "Department":
            donnees = donnees.sort_values("count", ascending=False)
            if len(donnees) > 20:
                top19 = donnees.iloc[:19]
                other_count = donnees["count"].iloc[19:].sum()
                donnees = pd.concat([top19, pd.DataFrame([["Other", other_count]], columns=["Department", "count"])], ignore_index=True)
            self.ax.pie(donnees["count"], labels=donnees["Department"], autopct='%1.1f%%', startangle=140)
            self.ax.set_title("Répartition par Département")
        elif question == "Q2":
            comptage = donnees["Réponse"].value_counts().sort_values(ascending=False)
            if len(comptage) > 20:
                top19 = comptage.iloc[:19]
                other_count = comptage.iloc[19:].sum()
                comptage = pd.concat([top19, pd.Series([other_count], index=["Other"])])
            self.ax.pie(comptage, labels=comptage.index, autopct='%1.1f%%', startangle=140)
            self.ax.set_title("Répartition des réponses Q2")
        else:
            comptage = donnees["Catégorie"].value_counts().sort_values(ascending=False)
            if len(comptage) > 20:
                top19 = comptage.iloc[:19]
                other_count = comptage.iloc[19:].sum()
                comptage = pd.concat([top19, pd.Series([other_count], index=["Other"])])
            self.ax.pie(comptage, labels=comptage.index, autopct='%1.1f%%', startangle=140)
            self.ax.set_title(f"Répartition des réponses {question}")

    # Méthode privée pour afficher un graphique en lignes
    def _afficher_line(self, donnees):
        question = self.var_question.get()
        if question == "Department":
            donnees = donnees.sort_values("count", ascending=False)
            if len(donnees) > 20:
                top19 = donnees.iloc[:19]
                other_count = donnees["count"].iloc[19:].sum()
                donnees = pd.concat([top19, pd.DataFrame([["Other", other_count]], columns=["Department", "count"])], ignore_index=True)
            self.ax.plot(donnees["Department"], donnees["count"], marker='o', color='green')
            self.ax.set_xlabel("Department")
            self.ax.set_ylabel("Nombre de répondants")
            self.ax.set_title("Tendance par Département")
            self.ax.tick_params(axis="x", rotation=45)
        elif question == "Q2":
            comptage = donnees["Réponse"].value_counts().sort_values(ascending=False)
            if len(comptage) > 20:
                top19 = comptage.iloc[:19]
                other_count = comptage.iloc[19:].sum()
                comptage = pd.concat([top19, pd.Series([other_count], index=["Other"])])
            self.ax.plot(comptage.index, comptage.values, marker='o', color='blue')
            self.ax.set_xlabel("Réponse")
            self.ax.set_ylabel("Mentions")
            self.ax.set_title("Tendance des mentions Q2")
            self.ax.tick_params(axis="x", rotation=45)
        else:
            comptage = donnees["Catégorie"].value_counts().sort_values(ascending=False)
            if len(comptage) > 20:
                top19 = comptage.iloc[:19]
                other_count = comptage.iloc[19:].sum()
                comptage = pd.concat([top19, pd.Series([other_count], index=["Other"])])
            self.ax.plot(comptage.index, comptage.values, marker='o', color='purple')
            self.ax.set_xlabel("Catégorie")
            self.ax.set_ylabel("Réponses 'Oui'")
            self.ax.set_title(f"Tendance des réponses {question}")
            self.ax.tick_params(axis="x", rotation=45)

    # Méthode pour gérer le zoom via la molette de la souris sur le graphique réseau
    def gestion_zoom(self, event):
        if self.type_visu_actuelle != "reseau" or self.var_question.get() != "Q2" or not event.inaxes:
            return
        facteur = self.ZOOM_IN_FACTOR if event.button == 'up' else self.ZOOM_OUT_FACTOR
        self.echelle_actuelle = max(self.MIN_ZOOM, min(self.echelle_actuelle * facteur, self.MAX_ZOOM))
        if self.G is not None:
            self._redessiner_reseau()

    # Méthode pour changer le type de visualisation et rafraîchir l'affichage
    def changer_visu(self, type_visu):
        if self.var_question.get() != "Q2" and type_visu in ("heatmap", "reseau"):
            type_visu = "barres"
        self.type_visu_actuelle = type_visu
        self.afficher_visualisation(type_visu)

    # Méthode pour actualiser l'affichage lors de la sélection d'une nouvelle question
    def actualiser_affichage(self, event=None):
        question = self.var_question.get()
        self.lbl_question.config(text=self.questions.get(question, ""))
        # Activation ou désactivation de certains boutons selon la question sélectionnée
        if question == "Q2":
            self.btn_heatmap.config(state=tk.NORMAL)
            self.btn_reseau.config(state=tk.NORMAL)
        else:
            self.btn_heatmap.config(state=tk.DISABLED)
            self.btn_reseau.config(state=tk.DISABLED)
        self.afficher_visualisation(self.type_visu_actuelle)

# Bloc principal pour démarrer l'application
if __name__ == "__main__":
    racine = tk.Tk()
    app = VisualisateurCSV(racine)
    racine.geometry("1200x800")
    racine.mainloop()