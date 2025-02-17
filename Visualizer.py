# Importation des modules nécessaires
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.patches import Patch
import seaborn as sns
import networkx as nx

# Classe principale pour visualiser les données CSV et générer divers graphiques
class VisualisateurCSV:
    def __init__(self, racine):
        # Initialisation de la fenêtre principale et des variables de l'application
        self.racine = racine
        self.racine.title("Visualisateur")
        self.df = None            # DataFrame qui contiendra les données CSV chargées
        self.G = None             # Graphe généré à partir des données (pour la visualisation en réseau)
        self.pos = None           # Positions des nœuds dans le graphe
        self.base_taille_noeud = 30   # Taille de base des nœuds dans le graphique réseau
        self.echelle_actuelle = 1.0    # Facteur d'échelle pour le zoom du réseau
        self.type_visu_actuelle = "barres"  # Type de visualisation par défaut
        self.liste_a_names = []   # Liste des valeurs de "A-Name" pour la coloration des nœuds
        
        # Dictionnaire contenant les questions à afficher selon la sélection de l'utilisateur
        self.questions = {
            "Q1": "AVEC QUEL DÉPARTEMENT AIMERAIS-TU TRAVAILLER?",
            "Q2": "AVEC QUI AIMERAIS-TU TRAVAILLER EN DEHORS DE TON DÉPARTEMENT?",
            "Q3": "AVEC QUEL DÉPARTEMENT AIMERAIS-TU PAS TRAVAILLER?",
            "Q4": "À TON AVIS, QUEL DÉPARTEMENT NE TROUVERAIT PAS D'INTÉRÊT PROFESSIONNEL À TRAVAILLER AVEC TOI?",
            "Department": "DÉPARTEMENTS QUI ONT LE PLUS RÉPONDU"
        }
        
        # Configuration de l'interface utilisateur
        self.configurer_interface()
        
    # Méthode pour configurer l'interface graphique
    def configurer_interface(self):
        # Création du panneau gauche (pour les contrôles et boutons)
        panneau_gauche = tk.Frame(self.racine, width=200, bg="#f0f0f0")
        panneau_gauche.pack(side=tk.LEFT, fill=tk.Y)
        
        # Création du panneau droit (pour l'affichage des graphiques)
        panneau_droit = tk.Frame(self.racine)
        panneau_droit.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH)
        
        # Label pour afficher la question sélectionnée
        self.lbl_question = tk.Label(panneau_droit, text="", font=('Arial', 10), wraplength=600)
        self.lbl_question.pack(pady=10)
        
        # Cadre pour afficher la légende (utilisé pour le graphique réseau)
        self.frame_legende = tk.Frame(panneau_droit)
        self.lbl_legende = tk.Label(self.frame_legende, text="", font=('Arial', 9))
        self.lbl_legende.pack()
        self.frame_legende.pack(pady=5)
        
        # Bouton pour charger un fichier CSV
        btn_charger = tk.Button(panneau_gauche, text="Charger CSV", command=self.charger_csv)
        btn_charger.pack(pady=10, padx=10, fill=tk.X)
        
        # Variable et combobox pour sélectionner la question à visualiser
        self.var_question = tk.StringVar()
        lbl_selection = tk.Label(panneau_gauche, text="Visualiser:", bg="#f0f0f0")
        lbl_selection.pack(pady=(20,5))
        self.combo_questions = ttk.Combobox(panneau_gauche, 
                                           textvariable=self.var_question,
                                           values=["Department", "Q1", "Q2", "Q3", "Q4"], 
                                           state="readonly")
        self.combo_questions.pack(padx=10, fill=tk.X)
        self.combo_questions.current(0)  # Sélection par défaut : "Department"
        self.combo_questions.bind("<<ComboboxSelected>>", self.actualiser_affichage)
        
        # Bouton pour afficher un graphique en barres
        self.btn_barres = tk.Button(panneau_gauche, text="Graphique en Barres", 
                                   command=lambda: self.changer_visu("barres"))
        self.btn_barres.pack(pady=5, padx=10, fill=tk.X)
        
        # Bouton pour afficher une matrice de chaleur
        self.btn_heatmap = tk.Button(panneau_gauche, text="Matrice de Chaleur", 
                                    command=lambda: self.changer_visu("heatmap"))
        self.btn_heatmap.pack(pady=5, padx=10, fill=tk.X)
        
        # Bouton pour afficher un réseau de relations
        self.btn_reseau = tk.Button(panneau_gauche, text="Réseau de Relations", 
                                   command=lambda: self.changer_visu("reseau"))
        self.btn_reseau.pack(pady=5, padx=10, fill=tk.X)
        
        # Création de la figure Matplotlib pour l'affichage des graphiques
        self.figure = plt.Figure(figsize=(10,7), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=panneau_droit)
        self.canvas.get_tk_widget().pack(expand=True, fill=tk.BOTH)
        
        # Barre d'outils de navigation Matplotlib
        self.toolbar = NavigationToolbar2Tk(self.canvas, panneau_droit)
        self.toolbar.update()
        
        # Connexion de l'événement de défilement (scroll) pour gérer le zoom
        self.canvas.mpl_connect("scroll_event", self.gestion_zoom)

    # Méthode pour charger un fichier CSV et mettre à jour l'affichage
    def charger_csv(self):
        # Ouverture d'une boîte de dialogue pour sélectionner un fichier CSV
        fichier = filedialog.askopenfilename(filetypes=[("Fichiers CSV", "*.csv")])
        if fichier:
            # Lecture du fichier CSV avec le séparateur ";" et stockage dans self.df
            self.df = pd.read_csv(fichier, sep=";")
            # Extraction des valeurs uniques de la colonne "A-Name" pour la visualisation du réseau
            self.liste_a_names = self.df["A-Name"].unique().tolist()
            # Actualisation de l'affichage avec la question actuellement sélectionnée
            self.actualiser_affichage()

    # Méthode pour préparer les données en fonction de la question sélectionnée
    def preparer_donnees(self):
        question = self.var_question.get()
        
        if question == "Department":
            # Comptage des occurrences pour chaque département
            return self.df["Department"].value_counts().reset_index()
        
        elif question == "Q2":
            # Création d'une liste des colonnes Q2 (de Q2-Name1 à Q2-Name9)
            colonnes = [f"Q2-Name{i}" for i in range(1,10)]
            # Transformation des colonnes en format long (melt) pour faciliter la visualisation
            df_fusion = self.df.melt(
                id_vars=["A-Name", "Department"],
                value_vars=colonnes,
                value_name="Réponse"
            )
            # Exclusion des réponses qui contiennent "----"
            return df_fusion[df_fusion["Réponse"] != "----"]
        
        else:
            # Pour Q1, Q3 et Q4, on sélectionne les colonnes qui commencent par le préfixe de la question
            prefixe = f"{question}-"
            colonnes = [col for col in self.df.columns if col.startswith(prefixe)]
            # Transformation en format long avec filtrage pour ne conserver que les réponses "Oui"
            df_fusion = self.df.melt(
                id_vars=["A-Name", "Department"],
                value_vars=colonnes,
                var_name="Catégorie",
                value_name="Réponse"
            )
            return df_fusion[df_fusion["Réponse"] == "Oui"]

    # Méthode pour afficher le graphique en fonction du type choisi
    def afficher_visualisation(self, type_visu):
        # Ne rien faire si aucun fichier CSV n'est chargé
        if self.df is None: 
            return
        
        # Effacer la figure précédente et créer un nouvel axe
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)
        question = self.var_question.get()
        # Préparation des données à afficher selon la question
        donnees = self.preparer_donnees()
        
        try:
            # Sélection du type de graphique à afficher
            if type_visu == "barres":
                self._afficher_barres(donnees, question)
            elif type_visu == "heatmap":
                self._afficher_heatmap(donnees)
            elif type_visu == "reseau":
                self._afficher_reseau(donnees)
                
            # Ajustement de la mise en page et mise à jour du canvas
            self.figure.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            # En cas d'erreur, affichage d'un message dans la console
            print(f"Erreur de visualisation : {str(e)}")

    # Méthode privée pour afficher un graphique en barres
    def _afficher_barres(self, donnees, question):
        if question == "Department":
            # Utilisation de seaborn pour tracer un graphique en barres (nombre de répondants par département)
            sns.barplot(x='count', y='Department', data=donnees, ax=self.ax, palette="viridis")
            self.ax.set_xlabel("Nombre de répondants")
        elif question == "Q2":
            # Comptage des réponses pour Q2 et affichage en barres (limité aux 20 premières valeurs)
            comptage = donnees["Réponse"].value_counts().head(20)
            comptage.plot(kind="bar", ax=self.ax, color="skyblue")
            self.ax.set_ylabel("Mentions")
        else:
            # Pour Q1, Q3 et Q4 : comptage des réponses "Oui" par catégorie et affichage en barres
            comptage = donnees["Catégorie"].value_counts()
            sns.barplot(x=comptage.values, y=comptage.index, ax=self.ax, palette="rocket")
            self.ax.set_ylabel("Réponses 'Oui'")
        
        # Configuration des axes pour afficher des nombres entiers et rotation des étiquettes de l'axe X
        self.ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
        plt.xticks(rotation=45, ha='right')

    # Méthode privée pour afficher une matrice de chaleur (heatmap)
    def _afficher_heatmap(self, donnees):
        if self.var_question.get() == "Q2":
            # Création d'une table de contingence entre "A-Name" et "Réponse"
            matrice = pd.crosstab(donnees["A-Name"], donnees["Réponse"])
            sns.heatmap(matrice, ax=self.ax, cmap="YlGnBu", cbar_kws={'label': 'Mentions'})

    # Méthode privée pour afficher un graphique en réseau
    def _afficher_reseau(self, donnees):
        # Création du graphe à partir des données (colonnes "A-Name" et "Réponse")
        self.G = nx.from_pandas_edgelist(donnees, "A-Name", "Réponse")
        # Calcul des positions des nœuds dans le graphe
        self.pos = nx.spring_layout(self.G, k=0.3)
        # Redessiner le graphe avec la méthode dédiée
        self._redessiner_reseau()

    # Méthode privée pour redessiner le graphique en réseau (utile lors du zoom)
    def _redessiner_reseau(self):
        # Effacer la figure et créer un nouvel axe
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)
        # Calcul de la taille des nœuds en fonction du facteur d'échelle actuel
        taille_noeud = self.base_taille_noeud * self.echelle_actuelle
        
        # Définition des couleurs pour chaque nœud en fonction de leur appartenance à la liste A-Name
        couleurs = []
        for node in self.G.nodes():
            if node in self.liste_a_names:
                couleurs.append("#1f77b4")  # Bleu pour A-Name
            else:
                couleurs.append("#2ca02c")  # Vert pour les autres
         
        # Dessin du graphe avec NetworkX
        nx.draw(self.G, self.pos, ax=self.ax,
                with_labels=True,
                node_size=taille_noeud,
                font_size=6,
                node_color=couleurs,
                edge_color="gray",
                width=0.5)
        
        # Création de la légende pour le graphique réseau
        legend_elements = [
            Patch(facecolor='#1f77b4', edgecolor='black', label='A-Name'),
            Patch(facecolor='#2ca02c', edgecolor='black', label='Autres noms')
        ]
        self.ax.legend(handles=legend_elements, loc='upper right')
        
        # Actualisation du canvas et mise à jour du label de légende
        self.canvas.draw()
        self.lbl_legende.config(text="Bleu - A-Name            -            Vert - Autres participants")

    # Méthode pour gérer le zoom via la molette de la souris sur le graphique réseau
    def gestion_zoom(self, event):
        # Le zoom s'applique uniquement pour la visualisation en réseau de Q2 et si l'événement se produit dans les axes
        if (self.type_visu_actuelle != "reseau" or 
            self.var_question.get() != "Q2" or 
            not event.inaxes):
            return
            
        # Déterminer le facteur de zoom en fonction de la direction de la molette (up ou down)
        facteur = 1.1 if event.button == 'up' else 0.9
        self.echelle_actuelle = max(0.5, min(self.echelle_actuelle * facteur, 5.0))
        
        # Redessiner le graphe avec la nouvelle échelle
        if self.G is not None:
            self._redessiner_reseau()

    # Méthode pour changer le type de visualisation et rafraîchir l'affichage
    def changer_visu(self, type_visu):
        self.type_visu_actuelle = type_visu
        self.afficher_visualisation(type_visu)

    # Méthode pour actualiser l'affichage lorsque la question sélectionnée change
    def actualiser_affichage(self, event=None):
        question = self.var_question.get()
        # Mettre à jour le texte de la question affichée
        self.lbl_question.config(text=self.questions.get(question, ""))
        
        # Définir l'état (activé ou désactivé) des boutons réseau et heatmap selon la question
        etat_reseau = tk.NORMAL if question == "Q2" else tk.DISABLED
        etat_heatmap = tk.NORMAL if question == "Q2" else tk.DISABLED
        
        self.btn_reseau.config(state=etat_reseau)
        self.btn_heatmap.config(state=etat_heatmap)
        
        # Afficher ou masquer le cadre de légende selon la question sélectionnée
        if question == "Q2":
            self.frame_legende.pack()
        else:
            self.frame_legende.pack_forget()
        
        # Afficher la visualisation par défaut en mode graphique en barres
        self.afficher_visualisation("barres")

# Bloc principal pour démarrer l'application
if __name__ == "__main__":
    racine = tk.Tk()
    app = VisualisateurCSV(racine)
    racine.geometry("1200x800")
    racine.mainloop()
