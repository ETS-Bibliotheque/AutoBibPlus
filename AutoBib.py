# © 2023 Benjamin Lepourtois <benjamin.lepourtois@gmail.com>
# Copyright: All rights reserved.
# See the license attached to the root of the project.

"""
Projet d'Automatisation de Rapports d'Analyses Bibliométriques:

Ce programme fait partie intégrante du projet de conception et développement d'outils automatisés pour la réalisation de rapports d'analyses bibliométriques.

Contexte:

  ● Stage de 12 semaines sur l'été 2023 (12 juin au 1er septembre) dans l'École de Technologie Supérieure, Montréal, Canada
  ● Mission principale:
Développer des outils permettant l'automatisation de certaines étapes de production de rapports d'analyses bibliométriques 
destinés à aider les chercheurs et chercheuses dans la planification de la mesure de l'impact de leurs contributions scientifiques.

Approche choisie:

Nous avons choisi d'utiliser un script Python pour gérer toute l'automatisation des rapports.
  ● Extraction des données: par les API des différentes plateformes utilisées (Scopus et SciVal) à l'aide de la bibliothèque publique "pybliometrics"
  ● Traitement des données: en Python à l'aide de la bibliothèque "pandas"
  ● Interface Homme-Machine: en QT avec une interface très simpliste basée sur une boîte de dialogue
  ● Exportation des données: en Python à l'aide de la bibliothèque "pywin32" vers un fichier "Workbook" MacroExcel (.xlsm)
  ● Mise en forme Excel: avec des routines VBA appelées par le script Python
  ● Réalisation du rapport Word: avec des routines VBA, appelées par le script Python, qui exportent les données et les graphiques réalisés sur un document Word
"""


import sys, os, ctypes, win32gui, subprocess, requests
from datetime import datetime
from pathlib import Path

# Utilisation de QT pour la création de l'Interface Homme-Machine (IHM ou HMI en anglais)
from PySide6.QtWidgets import QApplication, QMainWindow, QLineEdit, QVBoxLayout, QWidget, QPlainTextEdit, QMessageBox, QToolBar
from PySide6.QtGui import QFont, QIcon, QFontMetrics, QAction
from PySide6.QtCore import Qt, Slot

# Importations locales
from Include.Front import ExitBox, CustomTextEdit, LoadingDialog, AchievedMessageBox, ReconfigMessageBox, InfoAPI, Info, Timer

# Appliquer une feuille de style CSS pour le texte en couleur
text_style_warning = '"color: #D35230"'
text_style_parameter = '"color: #AA1C4F"'
text_style_question = '"color: #0C5E31"'

# Fonction permettant de checker la présence d'un fichier de configuration si non la création se fait alors (reprise de la bibliothèque "pybliometrics"), également fonction d'initialisation
def check_create_config(console: QPlainTextEdit, response: str, keys: list = None, first_time: bool = False):
    import configparser
    from Include.constants import CONFIG_FILE
    from Include.create_config import create_config

    # Read/create config file (with fixture for RTFD.io)
    config = configparser.ConfigParser()
    config.optionxform = str

    # Fichier de configuration existant ?
    if not CONFIG_FILE.exists():
        # Clef et token API déjà fournis ?
        if keys == None:
            # Affiche dans un premier temps le message de bienvenu (permet aussi de garder la couleur noire sur les textes et tableaux par défaut)
            if first_time:
                console.append("Bienvenue sur <b>AutoBib</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques de l'ÉTS!")
                console.append("")
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal:</p>'.format(text_style_parameter))
            # Rien ne se passe si la zone de texte rentrée est vide
            if response == '':
                return False, None
            # Sépare la zone de texte rentrée avec la virgule comme séparateur
            elements = response.split(',')
            # Check le nombre d'éléments et si le deuxième élément n'est pas nul
            if len(elements) < 2 or elements[1] == '':
                console.append('<p style={}>! Manque la clef et/ou le token</p>'.format(text_style_warning))
                console.append('')
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal:</p>'.format(text_style_parameter))
                return False, None
            # Suppression des espaces avant et après les éléments
            elements = [element.strip() for element in elements]
            # Test de la clef et du token par envoie d'une requête (API Scopus / AuthorRetrival sur Mohamed Cheriet)
            try:
                # Envoi de la requête à l'API Scopus
                response = requests.get("https://api.elsevier.com/content/author/author_id/56216876600", params={"apiKey": elements[0], "insttoken": elements[1], "httpAccept": "application/json"})
            except requests.exceptions.RequestException as e:
                error_message = "Une erreur s'est produite!\n\nSi l'erreur persiste, veuillez contacter le service technique de  votre établissement.\n\nDétails:\n" + str(e)
                error_dialog = QMessageBox(QMessageBox.Critical, "Erreur", error_message, QMessageBox.Ok)
                error_dialog.exec()
            
            # Vérification du statut de la réponse
            if response.status_code != 200:
                console.append('<p style={}>! Clef et/ou Token pour les API invalides (statut de la requête : {})</p>'.format(text_style_warning, response.status_code))
                console.append('')
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal:</p>'.format(text_style_parameter))
                return False, None
            
            # On peut passer à l'étape suivante : rentrer un chemin d'accès valide à un répertoire pour enregistrer les rapports par défaut
            console.append("\n\n")
            console.append("<p style={}>☼ Veuillez entrer le chemin d'accès ENTIER du répertoire où vous voulez enregistrer les rapports par défault (ex: C:\\Users\\Name\\Documents):</p>".format(text_style_parameter))
            return False, elements
        
        # Rien ne se passe si la zone de texte rentrée est vide
        if response == '':
            return False, keys
        
        # Supprimer les espaces de la zone de texte rentrée, en faire un vrai chemin d'accès et vérifier s'il existe en local sur la machine
        docs_path = Path(response.strip())
        if not docs_path.is_dir():
            console.append("<p style={}>! Chemin d'accès à un dossier non-valide</p>".format(text_style_warning))
            console.append('')
            console.append("<p style={}>☼ Veuillez entrer le chemin d'accès ENTIER du répertoire où vous voulez enregistrer les rapports par défault (ex: C:\\Users\\Name\\Documents):</p>".format(text_style_parameter))
            return False, keys
        
        # Suppression du fichier de configuration non-abouti puis création du fichier complet
        subprocess.run("del %userprofile%\.config\pybliometrics.cfg", shell=True)
        config = create_config(keys=[keys[0]], insttoken=keys[1], docs_path=response.strip())

        # Gestion de l'affichage de la zone de text (effacage puis affichage des diverses informations)
        console.setPlainText('')
        console.append(f"Le fichier de configuration a été créé avec succès au chemin d'accès : {CONFIG_FILE}.")
        console.append("Pour plus de détails, veuillez consulter https://pybliometrics.rtfd.io/en/stable/configuration.html.")
        console.append('')

    console.append("Bienvenue sur <b>AutoBib</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques de l'ÉTS!")
    console.append("\nLes commandes suivantes pourraient vous aider:\n\t- 0,[1;2]  puis <Entrée/Enter>: \tpermet de combiner des types sous le nom du 1er entre crochets, SEULEMENT pour les 2 types de publications\n\n\t- seulement <Entrée/Enter>: \t\tpermet de sélectionner les paramètres par défaut (pour les questions)\n\nLa barre d'outils en rouge peut être déplacée à l'aide de sa ligne de points à son extrémité.\n")
    console.append("<b>Tapez vos commandes dans la barre d'entrée de texte tout en bas de la page</b>, puis validez les en appuyant sur la touche &lt;Entrée/Enter&gt; de votre clavier.")
    console.append("\n")

    console.append('<span style="color: #0C5E31">● Veuillez entrer le nom et le prénom du chercheur [respectivement avec virgule comme séparateur]:</span>')  # Affiche la première interrogation dans la console
    return True, keys



# Créé la classe ConsoleWindow (classe fille de QMainWindow), qui est la fenêtre principale qui nous fait office de console/prompt 2.0
class ConsoleWindow(QMainWindow):
    # Définie le constructeur de la classe
    def __init__(self):
        super().__init__() # Permet de récupérer le constructeur de la classe mère: QMainWindow

        self.setWindowTitle("AutoBib: Logiciel d'automatisation des rapports d'analyses bibliométriques de l'ÉTS") # Définie le nom de la fenêtre
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/ETS_Logo.png"))  # Définit le logo

        # Définie la zone de texte (console)
        self.console = CustomTextEdit(self) # Instanciation

        # Définie la zone d'entrée de commandes de l'utilisateur
        self.input_box = QLineEdit(self) # Instanciation
        self.input_box.setStyleSheet("background-color: #DEDEDE; color: black; font-family: Consolas; font-size: 14pt; border: 2px solid black;") # Définie le CSS
        self.input_box.returnPressed.connect(self.handle_input) # Touche ENTER connectée à la méthode handle_input de la même classe

        # Définie la barre d'outils
        toolbar = QToolBar('Toolbar', self)
        # Définition du style QSS pour la barre d'outils
        toolbar.setStyleSheet("""
            QToolBar {
                border: 0px solid black;
            };

            background-color: #EF3635;
        """)
        self.addToolBar(Qt.TopToolBarArea, toolbar)

        # Ajout action retour
        actRedo = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Back.svg"), "&Retour", self)
        actRedo.setShortcut('Ctrl+Z')
        actRedo.triggered.connect(self.retour)
        toolbar.addAction(actRedo)

        # Ajout d'un séparateur
        toolbar.addSeparator()

        # Ajout action raz
        actReset = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Reset.svg"), "&Remise à zéro", self)
        actReset.setShortcut('Ctrl+R')
        actReset.triggered.connect(self.raz)
        toolbar.addAction(actReset)

        toolbar.addSeparator()

        # Ajout action feuille blanche
        actReset = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/WhiteSheet.svg"), "&Page blanche", self)
        actReset.setShortcut('Ctrl+N')
        actReset.triggered.connect(self.whitesheet)
        toolbar.addAction(actReset)
        
        toolbar.addSeparator()
        toolbar.addSeparator()

        # Ajout action reconfiguration
        actReconfig = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Config.svg"), "&Reconfiguration", self)
        actReconfig.setShortcut('F1')
        actReconfig.triggered.connect(self.reconfig)
        toolbar.addAction(actReconfig)

        toolbar.addSeparator()
        toolbar.addSeparator()

        # Ajout action Infos API
        actAPI = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/API.svg"), "&Infos API", self)
        actAPI.setShortcut('F2')
        actAPI.triggered.connect(self.API)
        toolbar.addAction(actAPI)

        toolbar.addSeparator()
        toolbar.addSeparator()

        # Ajout action Informations
        actInfos = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Infos.svg"), "&Informations", self)
        actInfos.setShortcut('F12')
        actInfos.triggered.connect(self.infos)
        toolbar.addAction(actInfos)

        # Définir la LoadingBox
        self.loading_dialog = LoadingDialog()

        # Organisation de la fenêtre Vertical Box et on met les deux éléments à la suite dans le bon ordre!
        layout = QVBoxLayout()
        layout.addWidget(self.console)
        layout.addWidget(self.input_box)

        # Créé la widget et ajoute le layout principal
        container = QWidget()
        container.setLayout(layout)
        container.setStyleSheet("background-color: #DEDEDE")  # Couleur d'arrière-plan de la widget principale

        # Set le widget sur la fenêtre principale
        self.setCentralWidget(container)

        # Initialisation
        self.text_style_question = '"color: #0C5E31"'
        self.input_box.setPlaceholderText("Entrez votre texte ici...")
        self.input_box.setFocus()
        # Instanciation du Timer
        self.timer = Timer()
        # Tableau pour les valeurs sur les clefs API
        self.infos_API = {'CitationOverview': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'}, \
                          'AuthorRetrieval': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'},\
                          'AuthorSearch': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'},\
                          'AuthorLookup': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'}}

        # Exécuter le fichier .bat de refresh du genpy pour une recherche plus rapide et éviter de futurs bugs liés aux fichiers de ce répertoire
        subprocess.run(os.path.dirname(os.path.abspath(__file__)) + "/maj_gen_py.bat", shell=True)

        # Backend:
        self.first_time = True # Permet de savoir si c'est la première fois que l'on rentre dans la fonction _affichageQuestions
        self.keys_valid = None
        # Variable pour stocker l'état courant de la machine à états
        validation, _ = check_create_config(self.console, '', first_time=True)
        self.state = 0 if validation else -1

        # Frontend
        self.df_doc_type_selected = [0,0]
        self.tableauQuestions = [
            '<span style={}>● Veuillez entrer le nom et le prénom du chercheur [respectivement avec virgule comme séparateur]:</span>'.format(self.text_style_question),
            "<span style={}>● Quel.s type.s de documents souhaitez-vous exclure de la liste? (Entrée: AUCUN)[index avec virgule comme séparateur]</span>".format(self.text_style_question),
            "<span style={}>● Quelles sont les 2 plages d'années que vous-choisissez? (Entrée: {}, {}, {} (soit: 3ans, 5ans, Carrière))[année avec virgule comme séparateur]</span>".format(self.text_style_question, 'NONE', 'NONE', 'NONE'),
            "<span style={}>● Quels sont les 2 types de publications que vous-voulez mettre en valeur? (Entrée: {}, {})[index avec virgule comme séparateur]</span>".format(self.text_style_question, 'NONE', 'NONE'),
            '<span style={}>● Pour une nouvelle recherche, veuillez taper "OTHER":</span>'.format(self.text_style_question)
        ]


    # Méthode qui réalise les différentes fonctions dès que l'utilisateur valide sa commande (appuie sur la touche "Entrée")
    def handle_input(self):
        # Stock la réponse de l'utilisateur, efface la zone d'entrée de texte et affiche la réponse sur la zone de texte
        self.response = self.input_box.text()
        self.input_box.clear()
        self.console.append('<span style="color: blue">{}</span>'.format("► " + self.response)) # Afficher l'entrée de l'utilisateur

        # Si l'état courant est à -1, la création du fichier de configuration se fait alors
        if self.state == -1:
            # Check fichier de configuration déjà créé ou non
            validation, keys = check_create_config(self.console, self.response, self.keys_valid)
            if not validation:
                self.keys_valid = keys
                return
            self.state = 0
            return
        
        from Include.Tools import homonyme, selection_homonyme, tab_graph_Collab, \
            selection_types_de_documents, donnees_documents_graph_citations, selection_plages_annees, tab_graph_citations, \
            tab_graph_publications, Excel_part1, Excel_part2, valeurs_encadre, selection_2_types_docs, tab_graph_SNIP
        from Include.author_search import AuthorSearch
        
        # Afficher le message de chargement
        self.loading_dialog.show()

        # Forcer le rafraîchissement de l'interface utilisateur
        QApplication.processEvents()

        # Obtient la largeur d'un caractère (qui est fixe car nous sommes en police monospace)
        font_metrics = QFontMetrics(QFont("Consolas", 11))
        width_char = font_metrics.averageCharWidth()

        # Test l'exécution du script de l'état courant et test la validité des transitions possibles sinon message d'erreur
        try :
            # Machine à états
            match self.state:
                # État initial : recherche d'un chercheur par son nom et son prénom
                case 0:
                    # Validation du format de la requête de l'utilisateur
                    if ',' in self.response and not self.response.split(",")[1] =='':
                        last_name = self.response.split(",")[0]
                        first_name = self.response.split(",")[1]
                    else:
                        self.console.append('<p style={}>! Manque du séparateur (virgule) et/ou du prénom du chercheur</p>'.format(text_style_warning))
                        self.console.append('')
                        self.console.append('<p style={}>● Veuillez entrer le nom et le prénom du chercheur [respectivement avec virgule comme séparateur]:</p>'.format(text_style_question))
                        # Fermer le message de chargement
                        self.loading_dialog.close()
                        return
                    
                    # Lancement du timer
                    self.timer.start()

                    # Recherche du chercheur
                    self.search = AuthorSearch('AUTHLAST(' + last_name + ') and AUTHFIRST(' + first_name + ')', refresh=True)
                    self.infos_API['AuthorSearch'].update({key: self.search._header[key] for key in self.search._header if key in self.infos_API['AuthorSearch']})
                    
                    # Incrémentation en fonction du nombre d'homonyme
                    self.state += homonyme(self.search, self.console, int(self.width()/width_char)-10)                

                    # S'il n'y a pas d'homonymes
                    self._rechercheSurChercheur() if self.state == 2 else None

                # État 1 : cas où la recherche a mené à des homonymes, il faut alors choisir l'un d'entre eux      
                case 1:
                    # Vérifie si le ou les numéros d'index sont correctes
                    if selection_homonyme(self.response, self.search, self.console):
                        self.state += 1
                        self._rechercheSurChercheur(choix=int(self.response))

                # État 2 : certains documents doivent être exclus et export des premières données vers le doc Excel
                case 2:
                    # Séparer les types de documents sélectionnés par l'utilisateur
                    selected_types = self.response.split(',')
                    selected_types = [element.strip() for element in selected_types]
    
                    # Vérifie la conformité de la commande de l'utilisateur
                    if len(selected_types) == 1 and selected_types[0] == "":
                        self.index_list = self.df_doc_type.index.tolist()
                    elif selection_types_de_documents(selected_types, len(self.df_doc_type), self.console):
                        # Obtenir une liste d'entier puis mettre à jour la liste
                        selected_types = [int(x) for x in selected_types]
                        self.index_list = [x for x in self.df_doc_type.index.tolist() if x not in selected_types]
                    else:
                        self._affichageQuestions(1)
                        return
                    
                    self.state += 1

                    # Calcul, mise en forme des données pour le graphique des citations
                    self.docs_list, self.selected_eids_list, self.years = donnees_documents_graph_citations(self.console, self.au_retrieval, self.df_doc_type, self.index_list)
                    self.df, self.nom_prenom, header = tab_graph_citations(self.console, self.au_retrieval, self.selected_eids_list, self.docs_list, int(self.width()/width_char)-10)
                    
                    # Mise à jour du dictionnaire des données sur les API
                    self.infos_API['CitationOverview'].update({key: header[key] for key in header if key in self.infos_API['CitationOverview']})
                    
                    # Calcul, mise en forme des données pour les valeurs de l'encadré
                    self.en_tete, self.annee_10y_adapt, header = valeurs_encadre(self.console, self.authorEID, self.years)
                    
                    # Mise à jour du dictionnaire des données sur les API
                    self.infos_API['AuthorLookup'].update({key: header[key] for key in header if key in self.infos_API['AuthorLookup']})

                    self.excel, self.classeur = Excel_part1(self.df, self.nom_prenom, self.en_tete, self.annee_10y_adapt)

                    self._affichageQuestions(2)

                # État 3 : les plages d'années des histogrammes sont choisies ici
                case 3:
                    validation, self.years_list, self.df_doc_type_selected = selection_plages_annees(self.console, self.response, self.years, self.df_doc_type, self.index_list)
                    if validation:
                        self.state += 1

                        self._affichageQuestions(3)
                    else:
                        self._affichageQuestions(2)

                # État 4 : les 2 types de publications mis en avant sur le graphique des Pubications se fait ici AINSI que l'envoie de toutes les autres données pour
                # le doc Excel ainsi que l'appel des routines VBA pour réaliser la mise en forme des données et la création de la fiche bibliométrique Word
                case 4:
                    validation, self.type_list = selection_2_types_docs(self.console, self.df_doc_type_selected, self.response)
                    if validation:
                        self.state = 0
                                            
                        self.df_pub = tab_graph_publications(self.console, self.au_retrieval, self.selected_eids_list, self.years_list, self.type_list, int(self.width()/width_char)-10)

                        self.years_list = [[int(item) for item in sublist] for sublist in self.years_list]
                        self.df_SNIP, header = tab_graph_SNIP(console=self.console, author_id=self.authorEID, years_list=self.years_list, window_width=int(self.width()/width_char)-10)
                        self.df_Collab, header = tab_graph_Collab(console=self.console, author_id=self.authorEID, years_list=self.years_list, window_width=int(self.width()/width_char)-10)
                        self.infos_API['AuthorLookup'].update({key: header[key] for key in header if key in self.infos_API['AuthorLookup']})

                        nom_classeur = self.classeur.Name
                        nom_classeur = nom_classeur.split('.')[0]
                        Excel_part2(self.excel, self.classeur, self.df_pub, self.df_SNIP, self.df_Collab)

                        # Fermer le message de chargement
                        self.loading_dialog.close()

                        self.timer.stop()
                        achieved_msg = AchievedMessageBox(time=self.timer.get_elapsed_time())
                        achieved_msg.exec()
                        self.timer.reset()

                        self.console.append("\n")
                        self.console.append("<b>Rapport d'analyse bibliométrique créé avec succès!</b>")
                        self.console.append("\n\n")

                        # Mettre la fenêtre Word en premier plan
                        try:
                            win32gui.SetForegroundWindow(win32gui.FindWindow(None, nom_classeur + " - Word"))
                        except:
                            win32gui.SetForegroundWindow(win32gui.FindWindow(None, nom_classeur + '.docx' + " - Word"))
                        
                        self._affichageQuestions(0)
                    else:
                        self._affichageQuestions(3)
        
        # Une erreur est rencontrée lors de l'exécution d'un des états, alors une boîte de dialogue s'affiche avec les détails de l'erreur
        except Exception as e:
            error_message = "Une erreur s'est produite!\n\nSi l'erreur persiste, veuillez contacter le service technique de  votre établissement.\n\nDétails:\n" + str(e)
            error_dialog = QMessageBox(QMessageBox.Critical, "Erreur", error_message, QMessageBox.Ok)
            error_dialog.exec()
            self.state = 0
            self._affichageQuestions(self.state)

        # Déplacer le QTextEdit à sa toute fin
        self.console.verticalScrollBar().setValue(self.console.verticalScrollBar().maximum())

        # Fermer le message de chargement
        self.loading_dialog.close()

    # Méthode : Gestion de l'affichage pour chaque état de la machine à états
    def _affichageQuestions(self, affichage_type: int):
        if affichage_type == 2:
            # Supprimer l'élément d'indice 2 (troisième élément) en utilisant pop()
            self.tableauQuestions.pop(2)
            self.tableauQuestions.insert(2, "<span style={}>● Quelles sont les 2 plages d'années que vous-choisissez? (Entrée: {}, {} (, {}) (soit: 3ans, 5ans, Carrière))[année avec virgule comme séparateur]</span>".format(self.text_style_question, datetime.now().year-3, datetime.now().year-5, self.years[0]))
        if affichage_type == 3:
            # Supprimer l'élément d'indice 3 (troisième élément) en utilisant pop()
            self.tableauQuestions.pop(3)
            liste_types_selec = self.df_doc_type_selected['Type de documents'].to_list()
            self.tableauQuestions.insert(3, "<span style={}>● Quels sont les 2 types de publications que vous-voulez mettre en valeur? (Entrée: {}, {} (soit: {}, {}))[index avec virgule comme séparateur]</span>".format(self.text_style_question, self.index_list[0], self.index_list[1] if len(self.index_list)>1 else '∅', liste_types_selec[0], liste_types_selec[1] if len(liste_types_selec)>1 else '∅')),
        self.console.append('')
        self.console.append(self.tableauQuestions[affichage_type])

    # Méthode : Recherche du chercheur sélectionné
    def _rechercheSurChercheur(self, choix: int = 0):
        from Include.Tools import retrieval, documents_selected_intro

        self.authorEID, self.au_retrieval = retrieval(choix, self.search, self.console)
        self.infos_API['AuthorRetrieval'].update({key: self.au_retrieval._header[key] for key in self.au_retrieval._header if key in self.infos_API['AuthorRetrieval']})

        self.df_doc_type = documents_selected_intro(self.console, self.au_retrieval)

        self._affichageQuestions(1)

    # Demande de fermeture de la fenêtre
    def closeEvent(self, event):
        message_box = ExitBox(self) # Instanciation
        message_box.exec()

        if message_box.clickedButton() == message_box.buttonYes:
            event.accept()
            self.timer.stop()
            self.timer.reset()
        else:
            event.ignore()


    @Slot()
    def retour(self):
        if self.state != 0 and self.state != -1:
            decalage = 1 if self.state > 2 else 0
            self.state -= 1 if self.state > 2 else self.state
            self._affichageQuestions(self.state - decalage)

    @Slot()
    def raz(self):
        if self.state != 0 and self.state != -1:
            self.state = 0
            self._affichageQuestions(self.state)

    @Slot()
    def reconfig(self):
        message_box = ReconfigMessageBox(self) # Instanciation
        message_box.exec()
        if self.state != -1 and message_box.clickedButton() == message_box.buttonYes:            
            self.state = -1
            self.console.setPlainText('')
            subprocess.run("del %userprofile%\\.config\\pybliometrics.cfg", shell=True)
            check_create_config(self.console, '', None, True)

    @Slot()
    def API(self):
        message_box = InfoAPI(self.infos_API) # Instanciation
        message_box.exec()

    @Slot()
    def infos(self):
        message_box = Info() # Instanciation
        message_box.exec()

    @Slot()
    def whitesheet(self):
        if self.state != -1:
            self.console.setPlainText('')
            self.console.append("Bienvenue sur <b>AutoBib</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques de l'ÉTS!")
            self.console.append("\nLes commandes suivantes pourraient vous aider:\n\t- 0,[1;2]  puis <Entrée/Enter>: \tpermet de combiner des types sous le nom du 1er entre crochets, SEULEMENT pour les 2 types de publications\n\n\t- seulement <Entrée/Enter>: \t\tpermet de sélectionner les paramètres par défaut (pour les questions)\n\nLa barre d'outils en rouge peut être déplacée à l'aide de sa ligne de points à son extrémité.\n")
            self.console.append("<b>Tapez vos commandes dans la barre d'entrée de texte tout en bas de la page</b>, puis validez les en appuyant sur la touche &lt;Entrée/Enter&gt; de votre clavier.")
            self.console.append("")

            self.state = 0
            self._affichageQuestions(self.state)

# Main loop
if __name__ == "__main__":
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ETS.Automatisation_Rapports_Bibliometriques') # Permet que l'OS voit l'exécution du script indépendante de Python et donc de changer le logo lors du script

    app = QApplication(sys.argv) # Instanciation d'une application QT avec les arguments système
    app.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/ETS_Logo.png")) # Affiche l'icône

    window = ConsoleWindow() # Instanciation de la fenêtre que nous venons de créer
    window.setStyleSheet("background-color: white; border: NONE;") # Couleur d'arrière-plan de la fenêtre

    window.showMaximized() # Affiche en plein écran

    sys.exit(app.exec()) # Pour la fermeture de l'application
