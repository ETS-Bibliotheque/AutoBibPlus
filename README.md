<!--
Author: Benjamin Lepourtois <benjamin.lepourtois@gmail.com>
Copyright: All rights reserved.
See the license attached to the root of the project.
-->


<h1 align="center">Projet d'Automatisation de Rapports d'Analyses Bibliométriques</h1>
<p align="center">Ce repository contient l'avancement du projet de conception et développement d'outils automatisés pour la réalisation de rapports d'analyses bibliométriques.
<br><br>English below!</p>

<h3><b>Contexte:</b></h3>
<p style="text-align: justify;">&emsp;<b> ● Stage de 12 semaines sur l'été 2023</b> (12 juin au 1er septembre) dans l'<b>École de Technologie Supérieure, Montréal, Canada</b>
<br>&emsp;<b> ● Mission principale: </b> <br>
Développer des outils permettant l'automatisation de certaines étapes de production de rapports d'analyses bibliométriques destinés à aider les chercheurs et chercheuses dans la planification de la mesure de l'impact de leurs contributions scientifiques. </p>

<h3><b>Approche choisie :</b></h3>
<p> Nous avons choisi d'utiliser <b>un script Python pour gérer toute l'automatisation</b> des rapports.
<br>&emsp; ● Extraction des données: par les API des différentes plateformes utilisées (Scopus et SciVal) à l'aide de la bibliothèque publique "pybliometrics"
<br>&emsp; ● Traitement des données: en Python à l'aide de la bibliothèque "pandas"
<br>&emsp; ● Interface Homme-Machine: en QT avec une interface très simpliste basée sur une boîte de dialogue
<br>&emsp; ● Exportation des données: en Python à l'aide de la bibliothèque "pywin32" vers un fichier "Workbook" MacroExcel (.xlsm)
<br>&emsp; ● Mise en forme Excel: avec des routines VBA appelées par le script Python
<br>&emsp; ● Réalisation du rapport Word: avec des routines VBA, appelées par le script Python, qui exportent les données et les graphiques réalisés sur un document Word

<br><b>Reprise de la bibliothèque publique de code :</b>
<br><b>pybliometrics (version 3.5.2) :</b> https://github.com/pybliometrics-dev/pybliometrics
<br><b>Article sur pybliometrics : </b> <a href="https://www.sciencedirect.com/science/article/pii/S2352711019300573">Rose, Michael E. and John R. Kitchin (2019): "pybliometrics: Scriptable bibliometrics using a Python interface to Scopus", SoftwareX 10 (2019) 100263</a></p>



<h1 align="center"><br><br>Automation of Bibliometric Analysis Reports</h1>
<p align="center">This repository contains the progress of the project to design and develop automated tools for producing bibliometric analysis reports.</p>

<h3><b>Context:</b></h3>
<p style="text-align: justify;">&emsp;<b> ● 12-week internship over the summer of 2023</b> (12 June to 1st September) at the <b>École de Technologie Supérieure, Montreal, Canada.</b>
<br>&emsp;<b> ● Main mission: </b> <br>
Develop tools to automate certain stages in the production of bibliometric analysis reports to help researchers understand the impact of their scientific contributions. </p>

<h3><b>Selected approach:</b></h3>
<p> We chose to use <b>a Python script to handle all the automation</b> of the reports.
<br>&emsp; ● Data extraction: via the APIs of the miscellaneous platforms used (Scopus and SciVal) using the "pybliometrics" public library
<br>&emsp; ● Data processing: in Python using the "pandas" library
<br>&emsp; ● Human Machine Interface: in QT with a very simplistic interface based on a dialog box
<br>&emsp; ● Data export: in Python using the "pywin32" library to a MacroExcel "Workbook" file (.xlsm)
<br>&emsp; ● Excel formatting: with VBA routines called by the Python script
<br>&emsp; ● Production of the Word report: with VBA routines, called by the Python script, which export the data and graphs produced to a Word document.

<br><b>Takeover of the code's public library:</b>
<br><b>pybliometrics (version 3.5.2):</b> https://github.com/pybliometrics-dev/pybliometrics
<br><b>Article on pybliometrics: </b> <a href="https://www.sciencedirect.com/science/article/pii/S2352711019300573">Rose, Michael E. and John R. Kitchin (2019): "pybliometrics: Scriptable bibliometrics using a Python interface to Scopus", SoftwareX 10 (2019) 100263</a></p>


<h2 align="left"><br>À propos de moi / About me:</h2>
<p>(2023/08/25) &emsp;English below!
<br>Je suis actuellement apprenti à l'École Centrale de Nantes (école d'ingénieur) et au sein du groupe français Thales en France (double statut: étudiant et salarié). 
<br>J'effectue ce stage, de première année de cycle d'ingénieur, dans le cadre de mon Projet de Séjour à l'International pour valider mon diplôme d'ingénieur.
<br><br>I'm currently an apprentice at the École Centrale de Nantes (engineering school) and with the French group Thales in France (dual status: student and employee). 
<br>I'm doing this work placement in my first year of engineering studies as part of my International Study Project to validate my engineering degree (equivalent to a master's degree).</p>

<h3 align="left"><br><br><br>Pour me contacter / Contact me:</h3>

<p align="left">
<a href="https://www.linkedin.com/in/benjamin-lepourtois-b09564232/" target="blank"><img align="center" src="https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/LinkedIn_icon.svg/2048px-LinkedIn_icon.svg.png" alt="v cvcv" height="30" width="30" /></a>
<link rel="stylesheet" href="Documentation/Stylesheet.css"/><link/>
</p>

<h3 align="left">Mes langages et mes outils de développeur / My developer languages and tools:</h3>
<div> 
    <p align="left"> 
        <a href="https://www.cprogramming.com/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/c/c-original.svg" alt="c" width="40" height="40"/> </a> 
        <a href="https://www.w3schools.com/cpp/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/cplusplus/cplusplus-original.svg" alt="cplusplus" width="40" height="40"/> </a> 
        <a href="https://www.python.org" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/python/python-original.svg" alt="python" width="40" height="40"/> </a> 
        <a href="https://www.qt.io/" target="_blank" rel="noreferrer"> <img src="https://upload.wikimedia.org/wikipedia/commons/0/0b/Qt_logo_2016.svg" alt="qt" width="40" height="40"/> </a> 
        <a href="https://git-scm.com/" target="_blank" rel="noreferrer"> <img src="https://www.vectorlogo.zone/logos/git-scm/git-scm-icon.svg" alt="git" width="40" height="40"/> </a> 
        <a href="https://www.jenkins.io" target="_blank" rel="noreferrer"> <img src="https://www.vectorlogo.zone/logos/jenkins/jenkins-icon.svg" alt="jenkins" width="40" height="40"/> </a> 
        <a href="https://www.linux.org/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/linux/linux-original.svg" alt="linux" width="40" height="40"/> </a> 
        <a href="https://www.w3.org/html/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/html5/html5-original-wordmark.svg" alt="html5" width="40" height="40"/> </a>  
        <a href="https://www.w3schools.com/css/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/css3/css3-original-wordmark.svg" alt="css3" width="40" height="40"/> </a> 
        <a href="https://www.docker.com/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/docker/docker-original-wordmark.svg" alt="docker" width="40" height="40"/> </a> 
        <a href="https://www.mathworks.com/" target="_blank" rel="noreferrer"> <img src="https://upload.wikimedia.org/wikipedia/commons/2/21/Matlab_Logo.png" alt="matlab" width="40" height="40"/> </a> 
        <a href="https://www.mysql.com/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/mysql/mysql-original-wordmark.svg" alt="mysql" width="40" height="40"/> </a> 
    </p> 
</div>

<div align="center" style="display: flex; justify-content: center; align-items: center;">
<br>
  <a href="https://www.etsmtl.ca/" style="margin-right: 50px;"> <img src="Logos/ETS_Logo.png" width="70" height="70"></a>
  <a href="https://www.ec-nantes.fr/" style="margin-left: 50px;"> <img src="Logos/CN.png" width="125" height="70"></a>
</div>
