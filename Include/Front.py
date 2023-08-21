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

import os, threading, time, typing, pytz

from datetime import datetime

from PySide6 import QtCore
# Utilisation de QT pour la création de l'Interface Homme-Machine (IHM ou HMI en anglais)
from PySide6.QtWidgets import QMessageBox, QTextEdit, QDialog, QVBoxLayout, QLabel, QApplication, QWidget, QPushButton, QPlainTextEdit, QTabWidget
from PySide6.QtGui import QFont, QIcon, QPixmap, QColor, QTextCursor, QCursor
from PySide6.QtCore import QTimer, Qt


# Redéfinition de la classe QMessageBox : QMessageBox customisée!
class CustomMessageBox(QMessageBox):
    def __init__(self, parent=None):
        super().__init__(parent)
              
        self.setStyleSheet("""
            QMessageBox { background-color: #DEDEDE; font-size: 12pt; }
            QLabel { background-color: #DEDEDE; }
            QPushButton { background-color: #EF3E45; color: #ffffff; font-size: 12pt; border-radius: 10px; padding: 8px; }
            QPushButton:hover { background-color: #ff5555; }
        """)

        self.setIcon(QMessageBox.Question)
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png"))
        self.setWindowTitle("Confirmation")
        self.setText("\nÊtes-vous sûr de vouloir quitter?")

        self.buttonYes = self.addButton("Oui", QMessageBox.YesRole)
        self.buttonYes.setFont(QFont("Arial", 11, QFont.Bold))
        self.buttonNo = self.addButton("Non", QMessageBox.NoRole)
        self.buttonNo.setFont(QFont("Arial", 11, QFont.Bold))

        pixmap = QPixmap(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/ETS_Logo.png")  # Chemin vers icône personnalisée
        self.setIconPixmap(pixmap.scaled(64, 64))  # Redimensionner l'icône et l'assigner


class CustomTextEdit(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setStyleSheet("background-color: #DEDEDE; color: black; font-family: Consolas; font-size: 11pt; border: NONE;") # Définie le CSS de la console
        self.setReadOnly(True) # En lecture seule
        self.setLineWrapMode(QTextEdit.WidgetWidth)  # Mode de retour à la ligne en fonction de la largeur du widget   WidgetWidth

    # def mousePressEvent(self, event):
        # Bloquer la sélection de texte en ignorant l'événement mousePressEvent
        # event.ignore()

    def focusOutEvent(self, event):
        # Lorsque le QTextEdit perd le focus, rétablir la couleur du texte par défaut
        self.moveCursor(QTextCursor.Start)
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())
        super().focusOutEvent(event)


class LoadingDialog(QDialog):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.label = QLabel("Exécution en cours, veuillez patienter...")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

        self.setWindowTitle("Chargement...")
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png"))
        self.setModal(True)

        # Afficher une icône de run pour le curseur de la souris
        self.setCursor(QCursor(Qt.WaitCursor))


class AchievedMessageBox(QMessageBox):
    def __init__(self, time: int = 0) -> None:
        super().__init__()

        self.setStyleSheet("""
            QMessageBox { background-color: #DEDEDE; font-size: 10pt; }
            QLabel { background-color: #DEDEDE; }
            QPushButton { background-color: #EF3E45; color: #ffffff; font-size: 12pt; border-radius: 10px; padding: 8px; }
            QPushButton:hover { background-color: #ff5555; }
        """)
        from .startup import DOCS_PATH
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png")) 
        self.setWindowTitle("Rapport réalisé")
        self.setText(f"Félicitations!\n\nRapport d'analyse bibliométrique créé avec succès.\n\nLe rapport est enregistré par défaut dans le répertoire suivant:\n{DOCS_PATH[0]}\n\nCe rapport a été généré en: {int(time/60)}min {time%60}s")

        pixmap = QPixmap(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/ETS_Logo.png")  # Chemin vers icône personnalisée
        self.setIconPixmap(pixmap.scaled(64, 64))  # Redimensionner l'icône et l'assigner
    

# Redéfinition de la classe QMessageBox : QMessageBox customisée pour la reconfiguration!
class ReconfigMessageBox(QMessageBox):
    def __init__(self, parent=None):
        super().__init__(parent)
              
        self.setStyleSheet("""
            QMessageBox { background-color: #DEDEDE; font-size: 12pt; }
            QLabel { background-color: #DEDEDE; }
            QPushButton { background-color: #EF3E45; color: #ffffff; font-size: 12pt; border-radius: 10px; padding: 8px; }
            QPushButton:hover { background-color: #ff5555; }
        """)

        self.setIcon(QMessageBox.Question)
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png"))
        self.setWindowTitle("Reconfiguration")
        self.setText("\nÊtes-vous sûr de vouloir reconfigurer\nla clef et le token institutionnel?")

        self.buttonYes = self.addButton("Oui", QMessageBox.YesRole)
        self.buttonYes.setFont(QFont("Arial", 11, QFont.Bold))
        self.buttonNo = self.addButton("Non", QMessageBox.NoRole)
        self.buttonNo.setFont(QFont("Arial", 11, QFont.Bold))

        pixmap = QPixmap(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/Config.svg")  # Chemin vers icône personnalisée
        self.setIconPixmap(pixmap.scaled(48, 48))  # Redimensionner l'icône et l'assigner
    

class InfoAPI(QDialog):
    def __init__(self, infos_API: dict):
        super().__init__()

        self.setWindowTitle("Informations sur les API")
        self.setGeometry(100, 100, 700, 250)

        layout = QVBoxLayout()
        tab_main = QTabWidget()

        tab_subScival = QTabWidget()

        tab = QWidget()
        tab_layout = QVBoxLayout()
        tab_layout.addWidget(self.create_info_widget(infos_API['AuthorLookup']))
        tab.setLayout(tab_layout)

        tab_subScival.addTab(tab, "AuthorLookup")
        tab_main.addTab(tab_subScival, "SciVal")


        tab_subScopus = QTabWidget()

        tab1 = QWidget()
        tab1_layout = QVBoxLayout()
        tab1_layout.addWidget(self.create_info_widget(infos_API['AuthorSearch']))
        tab1.setLayout(tab1_layout)

        tab2 = QWidget()
        tab2_layout = QVBoxLayout()
        tab2_layout.addWidget(self.create_info_widget(infos_API['AuthorRetrieval']))
        tab2.setLayout(tab2_layout)

        tab3 = QWidget()
        tab3_layout = QVBoxLayout()
        tab3_layout.addWidget(self.create_info_widget(infos_API['CitationOverview']))
        tab3.setLayout(tab3_layout)
        
        tab_subScopus.addTab(tab1, "AuthorSearch")
        tab_subScopus.addTab(tab2, "AuthorRetrieval")
        tab_subScopus.addTab(tab3, "CitationOverview")

        tab_main.addTab(tab_subScopus, "Scopus")
        layout.addWidget(tab_main)
        self.setLayout(layout)

    def create_info_widget(self, infos):
        info_widget = QPlainTextEdit()
        info_widget.setReadOnly(True)  # Pour empêcher l'édition du texte

        if infos['X-RateLimit-Limit'] != 'None':
            # Convertir le timestamp UNIX en objet datetime
            human_readable_date = datetime.fromtimestamp(int(infos['X-RateLimit-Reset']), tz=pytz.UTC)
            # Convertir la date en fuseau horaire local
            local_timezone = pytz.timezone('America/Montreal') 
            local_date = human_readable_date.astimezone(local_timezone)
            # Formater la date comme vous le souhaitez (par exemple, "Friday 18 August 2023 03:58:21")
            formatted_date = local_date.strftime('%A %d %B %Y %H:%M:%S')
        else:
            formatted_date = 'None'

        info_text = "Date:\t\t" + infos['Date'] + "\nX-RateLimit-Limit:\t" + infos['X-RateLimit-Limit'] + "\nX-RateLimit-Remaining:\t" + infos['X-RateLimit-Remaining'] + "\nX-RateLimit-Reset:\t" + infos['X-RateLimit-Reset'] + "\t(" + str(formatted_date) + " / à Montréal)\n\nLes valeurs sont indiquées que lorsque vous avez utilisé l'API concernée, depuis l'ouverture du logiciel."

        info_widget.setPlainText(info_text)
        return info_widget


class Info(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Informations sur AutoBib")
        layout = QVBoxLayout()

        info_widget = QTextEdit()
        info_widget.setReadOnly(True)  # Pour empêcher l'édition du texte
        
        info_widget.append("<b>© Benjamin Lepourtois (2023)</b>")
        info_widget.append("")
        info_widget.append("<b>Pour citer ou reprendre ce logiciel :</b>")
        info_widget.append("Benjamin Lepourtois. (2023). AutoBib. École de Technologie Supérieure. Sous licence CC BY.")

        layout.addWidget(info_widget)
        self.setLayout(layout)

class Timer:
    def __init__(self):
        self._start_time = None
        self._elapsed_time = 0
        self._timer_thread = None
        self._running = False
        self._stop_event = threading.Event()

    def _timer_function(self):
        while not self._stop_event.is_set():
            time.sleep(1)
            self._elapsed_time += 1

    def start(self):
        if not self._running:
            self._start_time = time.time()
            self._running = True
            self._stop_event.clear()
            self._timer_thread = threading.Thread(target=self._timer_function)
            self._timer_thread.start()

    def stop(self):
        if self._running:
            self._stop_event.set()
            self._timer_thread.join()
            self._running = False

    def reset(self):
        self._start_time = None
        self._elapsed_time = 0

    def get_elapsed_time(self):
        if self._start_time is None:
            return 0
        return self._elapsed_time + int(time.time() - self._start_time)
    

