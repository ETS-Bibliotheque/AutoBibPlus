AutoBib :

Projet d'Automatisation de Rapports d'Analyses Bibliométriques

Ce repository contient l'avancement du projet de conception et développement d'outils automatisés pour la réalisation de rapports d'analyses bibliométriques.

English below!


Contexte:

  ● Stage de 12 semaines sur l'été 2023 (12 juin au 1er septembre) dans l'École de Technologie Supérieure, Montréal, Canada
  ● Mission principale:
Développer des outils permettant l'automatisation de certaines étapes de production de rapports d'analyses bibliométriques destinés à aider les chercheurs et chercheuses dans la planification de la mesure de l'impact de leurs contributions scientifiques.
Approche choisie :

Nous avons choisi d'utiliser un script Python pour gérer toute l'automatisation des rapports.
  ● Extraction des données: par les API des différentes plateformes utilisées (Scopus et SciVal) à l'aide de la bibliothèque publique "pybliometrics"
  ● Traitement des données: en Python à l'aide de la bibliothèque "pandas"
  ● Interface Homme-Machine: en QT avec une interface très simpliste basée sur une boîte de dialogue
  ● Exportation des données: en Python à l'aide de la bibliothèque "pywin32" vers un fichier "Workbook" MacroExcel (.xlsm)
  ● Mise en forme Excel: avec des routines VBA appelées par le script Python
  ● Réalisation du rapport Word: avec des routines VBA, appelées par le script Python, qui exportent les données et les graphiques réalisés sur un document Word


Reprise de la bibliothèque publique de code :
pybliometrics (version 3.5.2) : https://github.com/pybliometrics-dev/pybliometrics
Article sur pybliometrics : Rose, Michael E. and John R. Kitchin (2019): "pybliometrics: Scriptable bibliometrics using a Python interface to Scopus", SoftwareX 10 (2019) 100263





Automation of Bibliometric Analysis Reports

This repository contains the progress of the project to design and develop automated tools for producing bibliometric analysis reports.
Context:

  ● 12-week internship over the summer of 2023 (12 June to 1st September) at the École de Technologie Supérieure, Montreal, Canada.
  ● Main mission:
Develop tools to automate certain stages in the production of bibliometric analysis reports to help researchers understand the impact of their scientific contributions.
Selected approach:

We chose to use a Python script to handle all the automation of the reports.
  ● Data extraction: via the APIs of the miscellaneous platforms used (Scopus and SciVal) using the "pybliometrics" public library
  ● Data processing: in Python using the "pandas" library
  ● Human Machine Interface: in QT with a very simplistic interface based on a dialog box
  ● Data export: in Python using the "pywin32" library to a MacroExcel "Workbook" file (.xlsm)
  ● Excel formatting: with VBA routines called by the Python script
  ● Production of the Word report: with VBA routines, called by the Python script, which export the data and graphs produced to a Word document.


Takeover of the code's public library:
pybliometrics (version 3.5.2): https://github.com/pybliometrics-dev/pybliometrics
Article on pybliometrics: Rose, Michael E. and John R. Kitchin (2019): "pybliometrics: Scriptable bibliometrics using a Python interface to Scopus", SoftwareX 10 (2019) 100263





À propos de moi / About me (Benjamin Lepourtois) :

(2023/08/25)  English below!
Je suis actuellement apprenti à l'École Centrale de Nantes (école d'ingénieur) et au sein du groupe français Thales en France (double statut: étudiant et salarié).
J'effectue ce stage, de première année de cycle d'ingénieur, dans le cadre de mon Projet de Séjour à l'International pour valider mon diplôme d'ingénieur.

I'm currently an apprentice at the École Centrale de Nantes (engineering school) and with the French group Thales in France (dual status: student and employee).
I'm doing this work placement in my first year of engineering studies as part of my International Study Project to validate my engineering degree (equivalent to a master's degree).