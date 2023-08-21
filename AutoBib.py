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


import sys, os, ctypes, time, win32gui, subprocess, requests
from datetime import datetime
from pathlib import Path

# Utilisation de QT pour la création de l'Interface Homme-Machine (IHM ou HMI en anglais)
from PySide6.QtWidgets import QApplication, QMainWindow, QLineEdit, QVBoxLayout, QWidget, QPlainTextEdit, QMessageBox, QToolBar
from PySide6.QtGui import QFont, QIcon, QFontMetrics, QAction
from PySide6.QtCore import Qt, Slot

# Importations locales
# sys.path.append(os.path.dirname(os.path.abspath(__file__)) + '\Include')
from Include.Front import CustomMessageBox, CustomTextEdit, LoadingDialog, AchievedMessageBox, ReconfigMessageBox, InfoAPI, Info, Timer

# Appliquer une feuille de style CSS pour le texte en couleur
text_style_warning = '"color: #D35230"'
text_style_parameter = '"color: #AA1C4F"'


def check_create_config(console: QPlainTextEdit, response: str, keys: list = None, first_time: bool = False):
    import configparser
    from Include.constants import CONFIG_FILE
    from Include.create_config import create_config

    # Read/create config file (with fixture for RTFD.io)
    config = configparser.ConfigParser()
    config.optionxform = str

    if not CONFIG_FILE.exists():
        if keys == None:
            if first_time:
                console.append("Bienvenue sur <b>AutoBib</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques de l'ÉTS!")
                console.append("")
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal:</p>'.format(text_style_parameter))
            if response == '':
                return False, None
            elements = response.split(',')
            if len(elements) < 2 or elements[1] == '':
                console.append('<p style={}>! Manque la clef et/ou le token</p>'.format(text_style_warning))
                console.append('')
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal:</p>'.format(text_style_parameter))
                return False, None
            elements = [element.strip() for element in elements]
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
            
            console.append("\n\n")
            console.append("<p style={}>☼ Veuillez entrer le chemin d'accès ENTIER du répertoire où vous voulez enregistrer les rapports par défault (ex: C:\\Users\\Name\\Documents):</p>".format(text_style_parameter))
            return False, elements
        
        if response == '':
            return False, keys
        
        docs_path = Path(response.strip())
        if not docs_path.is_dir():
            console.append("<p style={}>! Chemin d'accès à un dossier non-valide</p>".format(text_style_warning))
            console.append('')
            console.append("<p style={}>☼ Veuillez entrer le chemin d'accès ENTIER du répertoire où vous voulez enregistrer les rapports par défault (ex: C:\\Users\\Name\\Documents):</p>".format(text_style_parameter))
            return False, keys
        
        subprocess.run("del %userprofile%\.config\pybliometrics.cfg", shell=True)
        config = create_config(keys=[keys[0]], insttoken=keys[1], docs_path=response.strip())

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
        # self.setGeometry(100, 100, 800, 600) # Définit la position et la taille par défaut de la fenêtre
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

        # Exécuter le fichier .bat de refresh du genpy pour une recherche plus rapide
        subprocess.run(os.path.dirname(os.path.abspath(__file__)) + "/maj_gen_py.bat", shell=True)

        # Backend:
        self.first_time = True # Permet de savoir si c'est la première fois que l'on rentre dans la fonction affichageQuestions
        self.keys_valid = None
        # Variable pour stocker l'état courant de la machine à états
        validation, keys = check_create_config(self.console, '', first_time=True)
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


    # Méthode qui réalise les différentes fonctions dès que l'utilisateur valide sa commande
    def handle_input(self):
        self.response = self.input_box.text()
        self.input_box.clear()
        self.console.append('<span style="color: blue">{}</span>'.format("► " + self.response)) # Afficher l'entrée de l'utilisateur

        # # Afficher une icône de run pour le curseur de la souris
        # self.setCursor(QCursor(Qt.WaitCursor))

        if self.state == -1:
            # Check fichier de configuration déjà créé ou non
            validation, keys = check_create_config(self.console, self.response, self.keys_valid)
            if not validation:
                self.keys_valid = keys
                return
            self.state = 0
            return
        
        from Include.Tools import rechercheChercheur, selectionEID, recupEID, documents_selected_intro, \
            documents_part_selection, documents_part_final, nb_publications_annees_intro, data_main, \
            nb_publications, Excel_part1, Excel_part2, vals_entete, documents_part_selection2, vals_SNIP, \
            vals_Collab
        
        
        # Afficher le message de chargement
        self.loading_dialog.show()

        # Forcer le rafraîchissement de l'interface utilisateur
        QApplication.processEvents()

        # Obtient la largeur d'un caractère (qui est fixe car nous sommes en police monospace)
        font_metrics = QFontMetrics(QFont("Consolas", 11))
        width_char = font_metrics.averageCharWidth()

        try :
            # Machine à états
            match self.state:
                case 0:
                    self.timer.start()
                    self.search, self.choix = rechercheChercheur(self.console, self.response, int(self.width()/width_char)-10)
                    self.infos_API['AuthorSearch'].update({key: self.search._header[key] for key in self.search._header if key in self.infos_API['AuthorSearch']})

                    if not self.choix:
                        self.state += 2
                        self.authorEID, self.au_retrieval = recupEID(self.console, self.search, 0)
                        self.infos_API['AuthorRetrieval'].update({key: self.au_retrieval._header[key] for key in self.au_retrieval._header if key in self.infos_API['AuthorRetrieval']})

                        self.df_doc_type = documents_selected_intro(self.console, self.au_retrieval)

                        self.affichageQuestions(1)
                    elif self.choix == 1:
                        self.state += 1
                        
                case 1:
                    if selectionEID(self.console, self.search, self.response):
                        self.state += 1
                        self.authorEID, self.au_retrieval = recupEID(self.console, self.search, int(self.response))
                        self.infos_API['AuthorRetrieval'].update({key: self.au_retrieval._header[key] for key in self.au_retrieval._header if key in self.infos_API['AuthorRetrieval']})
                        
                        self.df_doc_type = documents_selected_intro(self.console, self.au_retrieval)
                        
                        self.affichageQuestions(1)

                case 2:
                    validation, self.index_list = documents_part_selection(self.console, self.df_doc_type, self.response)
                    if validation:
                        self.state += 1
                        self.docs_list, self.selected_eids_list, self.years = documents_part_final(self.console, self.au_retrieval, self.df_doc_type, self.index_list)
                        # print(self.years)
                        self.df, self.nom_prenom, header = data_main(self.console, self.au_retrieval, self.selected_eids_list, self.docs_list, int(self.width()/width_char)-10)
                        self.infos_API['CitationOverview'].update({key: header[key] for key in header if key in self.infos_API['CitationOverview']})
                        
                        self.en_tete, self.annee_10y_adapt, header = vals_entete(self.console, self.authorEID, self.years)
                        self.infos_API['AuthorLookup'].update({key: header[key] for key in header if key in self.infos_API['AuthorLookup']})

                        self.excel, self.classeur = Excel_part1(self.df, self.nom_prenom, self.en_tete, self.annee_10y_adapt)

                        self.affichageQuestions(2)
                    else:
                        self.affichageQuestions(1)

                case 3:
                    validation, self.years_list, self.df_doc_type_selected = nb_publications_annees_intro(self.console, self.response, self.years, self.df_doc_type, self.index_list)
                    if validation:
                        self.state += 1

                        self.affichageQuestions(3)
                    else:
                        self.affichageQuestions(2)

                case 4:
                    validation, self.type_list = documents_part_selection2(self.console, self.df_doc_type_selected, self.response)
                    if validation:
                        self.state = 0
                                            
                        self.df_pub = nb_publications(self.console, self.au_retrieval, self.selected_eids_list, self.years_list, self.type_list, int(self.width()/width_char)-10)

                        self.years_list = [[int(item) for item in sublist] for sublist in self.years_list]
                        self.df_SNIP, header = vals_SNIP(console=self.console, author_id=self.authorEID, years_list=self.years_list, window_width=int(self.width()/width_char)-10)
                        self.df_Collab, header = vals_Collab(console=self.console, author_id=self.authorEID, years_list=self.years_list, window_width=int(self.width()/width_char)-10)
                        self.infos_API['AuthorLookup'].update({key: header[key] for key in header if key in self.infos_API['AuthorLookup']})

                        nom_classeur = self.classeur.Name
                        nom_classeur = nom_classeur.split('.')[0]
                        Excel_part2(self.excel, self.classeur, self.df_pub, self.df_SNIP, self.df_Collab)

                        # Fermer le message de chargement
                        self.loading_dialog.close()

                        # self.activateWindow() # Mettre en premier plan l'application QT
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
                        
                        self.affichageQuestions(0)
                    else:
                        self.affichageQuestions(3)
        
        except Exception as e:
            error_message = "Une erreur s'est produite!\n\nSi l'erreur persiste, veuillez contacter le service technique de  votre établissement.\n\nDétails:\n" + str(e)
            error_dialog = QMessageBox(QMessageBox.Critical, "Erreur", error_message, QMessageBox.Ok)
            error_dialog.exec()
            self.state = 0
            self.affichageQuestions(self.state)

        # Déplacer le QTextEdit à sa toute fin
        self.console.verticalScrollBar().setValue(self.console.verticalScrollBar().maximum())

        # Fermer le message de chargement
        self.loading_dialog.close()

        # # Restaurer l'icône du curseur de la souris à son état par défaut
        # QApplication.restoreOverrideCursor()

    # Gestion de l'affichage pour chaque état de la machine à états
    def affichageQuestions(self, affichage_type: int):
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

    # Demande de fermeture de la fenêtre
    def closeEvent(self, event):
        message_box = CustomMessageBox(self) # Instanciation
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
            self.affichageQuestions(self.state - decalage)

    @Slot()
    def raz(self):
        if self.state != 0 and self.state != -1:
            self.state = 0
            self.affichageQuestions(self.state)

    @Slot()
    def reconfig(self):
        message_box = ReconfigMessageBox(self) # Instanciation
        message_box.exec()
        if self.state != -1 and message_box.clickedButton() == message_box.buttonYes:            
            self.state = -1
            self.console.setPlainText('')
            subprocess.run("del %userprofile%\.config\pybliometrics.cfg", shell=True)
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
            self.affichageQuestions(self.state)

# Main loop
if __name__ == "__main__":
    myappid = 'ETS.Automatisation_Rapports_Bibliometriques'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) # Permet que l'OS voit l'exécution du script indépendante de Python et donc de changer le logo lors du script

    app = QApplication(sys.argv) # Instanciation d'une application QT avec les arguments système
    app.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/ETS_Logo.png")) # Affiche l'icône

    window = ConsoleWindow() # Instanciation de la fenêtre que nous venons de créer
    window.setStyleSheet("background-color: white; border: NONE;") # Couleur d'arrière-plan de la fenêtre

    window.showMaximized() # Affiche en plein écran

    sys.exit(app.exec()) # Pour la fermeture de l'application
