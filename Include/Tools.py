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

import os, unicodedata, win32gui, time, re
import pandas as pd
import win32com.client as win32

from datetime import datetime
from .abstract_citation import CitationOverview
from .author_retrieval import AuthorRetrieval
from .author_search import AuthorSearch
from .author_lookup import AuthorLookup
from .startup import DOCS_PATH

from PySide6.QtWidgets import QPlainTextEdit

# Appliquer une feuille de style CSS pour le texte en couleur
text_style_warning = '"color: #D35230"'
text_style_question = '"color: #0C5E31"'

# Dictionnaire des traductions souhaitées
trad_en2fr = {
    'Article': 'Article',
    'Review': 'Article de synthèse',
    'Conference Paper': 'Conférence',
    'Conference Review': 'Synthèse de conférence',
    'Data Paper': 'Article de données',
    'Editorial': 'Éditorial',
    'Book': 'Livre',
    'Book Chapter': 'Chapitre',
    'Erratum': 'Erratum',
    'Note': 'Commentaire',
    'Letter': "Lettre d'opinion",
    'Short Survey': 'Enquête',
    'Retracted': 'Rétractation',
    '∅': '∅'
}

# Inverser le dictionnaire en échangeant les clés et les valeurs
trad_fr2en = {v: k for k, v in trad_en2fr.items()}

# Fonction qui retourne l'incrément de la variable d'état en fonction du nombre de résultats pour un nom et un prénom de chercheur
def homonyme(resultatRecherche: AuthorSearch, console: QPlainTextEdit, window_width: int):
    if not resultatRecherche.get_results_size():
        console.append('<p style={}>! Aucun résultat</p>'.format(text_style_warning))
        console.append('')
        console.append('<p style={}>● Veuillez entrer le nom et le prénom du chercheur [respectivement avec virgule comme séparateur]:</p>'.format(text_style_question))
        return 0
    else:
        # Affiche les résultats de manière organisée
        pd.set_option('display.max_columns', None)
        df = pd.DataFrame(resultatRecherche.authors)
        df.index.name = 'Index'
        df = df.drop(df.columns[[0, 1, 3, 7]], axis=1)
        df.columns = ['Nom', 'Prénom', 'Affiliation', 'Nb docs', 'Ville', 'Pays', 'Domaine(s) de recherche']
        console.append('')
        console.append('<p style="text-decoration: underline; color: black;">Chercheur.s trouvé.s:</p>')
        console.append(df.to_string(index=True, col_space=0, line_width=window_width)) # .to_string().encode('utf-8')
        # Permet de savoir si l'utilisateur doit choisir un chercheur dans une liste
        if len(resultatRecherche.authors) > 1:
            console.append('')
            console.append('<p style={}>● Quel chercheur choisissez-vous (index)?</p>'.format(text_style_question))
            return 1
    return 2

# Fonction qui retourne vrai si les index rentrés sont valides
def selection_homonyme(choix: str, s: AuthorSearch, console: QPlainTextEdit):
    selection = False
    
    # Vérifier si la valeur entrée est un entier
    if choix.isdigit():          
        # Vérifier si la valeur entrée est un index valide
        if 0 <= int(choix) < len(s.authors):
            selection = True
        else:
            console.append('<p style={}>! Veuillez entrer un index du tableau valide</p>'.format(text_style_warning))
            console.append('')
            console.append('<p style={}>● Quel chercheur choisissez-vous (index)?</p>'.format(text_style_question))
    else:
        console.append('<p style={}>! Veuillez entrer un entier</p>'.format(text_style_warning))
        console.append('')
        console.append('<p style={}>● Quel chercheur choisissez-vous (index)?</p>'.format(text_style_question))

    return selection



# Fonction qui retourne l'EID tronqué et surtout l'instance de AuthorRetrieval sur le chercheur sélectionné
def retrieval(choix: int, s: AuthorSearch, console: QPlainTextEdit):
    # Récupération de l'identifier de l'eid en fonction du chercheur sélectionné
    author_eid = s.authors[choix].eid
    author_eid = author_eid.split("s2.0-")[-1] # récupère le 2ème élément créé avec le split (donc eid)

    # Affiche un résumé sur le chercheur sélectionné
    au_retrieval = AuthorRetrieval(author_eid, refresh=True)
    sum = str(au_retrieval)

    console.append('')
    console.append('<p style="text-decoration: underline;">Résumé du chercheur sélectionné:</p>')
    console.append('<p style="font-weight: bold;">{}</p>'.format(sum))
    console.append('')

    return author_eid, au_retrieval

# Fonction qui retourne un DataFrame sur les types de documents avec leur nombre en fonction du chercheur sélectionné
def documents_selected_intro(console: QPlainTextEdit, au_retrieval: AuthorRetrieval):
    docs = pd.DataFrame(au_retrieval.get_documents(refresh=10))

    # Afficher le nombre de valeurs différentes
    nb_val = docs['subtypeDescription'].nunique()

    # Afficher les valeurs uniques dans la colonne 'subtypeDescription'
    list_val = docs['subtypeDescription'].unique()

    total = 0
    for i in range(0, nb_val):
        compte_valeurs = len(docs[docs['subtypeDescription'] == list_val[i]])
        total += compte_valeurs

    # Compter les occurrences de chaque valeur
    value_counts = docs['subtypeDescription'].value_counts()

    # Créer un DataFrame avec index de ref et données={type_de_document, value_counts}
    df = pd.DataFrame({'count': value_counts})
    df = df.reset_index()
    df.columns = ['Type de documents', 'Nombre']
    df.index.name = 'Index'

    # Appliquer la traduction aux types de documents
    df['Type de documents'] = df['Type de documents'].map(trad_en2fr)

    console.append('<p style="text-decoration: underline; color: black;">Nb de documents en fonction de leur type:</p>')
    console.append(df.to_string(index=True, col_space=0, line_width=200)) # .to_string().encode('utf-8')
    console.append('<p style="font-weight: bold;">Total: {}</p>'.format(str(total)))

    return df



# Fonction qui retourne vrai si la sélection des types est correcte
def selection_types_de_documents(selected_types: list, len_df: int, console: QPlainTextEdit):
    tout_valide = True

    for type_index in selected_types:
        # Vérifier si l'index est valide
        if not (type_index.isdigit() and int(type_index) < len_df):
            console.append('<p style={}>! Index non valide: {}</p>'.format(text_style_warning, type_index))
            tout_valide = False
            continue
        elif selected_types.count(str(int(type_index))) > 1:
            console.append('<p style={}>! Doublon trouvé: {}</p>'.format(text_style_warning, type_index))
            tout_valide = False        
    
    return tout_valide

# Fonction qui retourne les listes de : du nombre de documents par année avec prise en compte des types de docs sélectionnés, 
# des eids de tous les documents des types sélectionnés, ainsi que les années de carrière du chercheur
def donnees_documents_graph_citations(console: QPlainTextEdit, au_retrieval: AuthorRetrieval, df: pd.DataFrame, selected_types: list):
    # Créé un DataFrame avec toutes les données sur tous les documents du chercheur sélectionné
    docs = pd.DataFrame(au_retrieval.get_documents(refresh=10))
    # Trie par année de publication des documents
    draft_list = docs['coverDate'].str[:4].sort_values() # pour draft list 
    # Liste tous les types sélectionnés avec la retraduction en anglais
    selected_types = [trad_fr2en[doc] for doc in [df['Type de documents'].loc[int(type_index)] for type_index in selected_types]]

    # Filtrer les documents en fonction des types sélectionnés
    filtered_docs = [len(docs[docs['subtypeDescription'] == type_info]) for type_info in selected_types]

    # Créer un DataFrame avec index=list_val et données=value_counts
    df2 = pd.DataFrame({'Type de documents': selected_types,
                        'Nombre': filtered_docs})

    selection_string = ', '.join([trad_en2fr[doc] for doc in df2['Type de documents'].tolist()])
    console.append('')
    console.append('<a style="font-weight: bold;">Votre sélection: </a>' + selection_string)
    console.append('')

    # Filtrer les documents en fonction des types sélectionnés et créer une nouvelle colonne 'Année'
    filtered_docs2 = docs.loc[docs['subtypeDescription'].isin(selected_types), ['subtypeDescription', 'coverDate']].copy()
    filtered_docs2['Année'] = filtered_docs2['coverDate'].str[:4]

    # Grouper les documents filtrés par année et effectuer la somme des articles par année
    counts_par_annee = filtered_docs2.groupby('Année').size()

    # Créer un DataFrame à partir des comptages par année et affichage avec transposition
    df_final = pd.DataFrame({'Nombre': counts_par_annee})

    # Créé la liste utilisée pour l'exportation des données dans Excel
    list = df_final.reset_index().values.tolist()
    first_year = draft_list.values[0]
    total_annees = datetime.now().year - int(first_year) + 2

    # Liste de toutes les années du chercheur
    years = [int(first_year) + i for i in range(total_annees)]

    final_list = [0] * total_annees
    decount = 0
    for i in range(0, total_annees):
        if i - decount < len(list):
            if int(list[i - decount][0]) == int(first_year) + i:
                final_list[i] = list[i - decount][1]
                continue
        decount += 1
        final_list[i] = 0

    final_list.append(sum(final_list))

    eids_list = docs.loc[docs['subtypeDescription'].isin(selected_types), ['eid']].copy()
    eids_list = eids_list['eid'].tolist()

    return final_list, eids_list, years

# Fonction qui retourne les listes de : du nombre de citations par année et les années de carrière du chercheur
def donnees_citations_graph_citations(au_retrieval: AuthorRetrieval, document_eids: list):
    first_year = au_retrieval.publication_range[0]
    total_annees = datetime.now().year - first_year + 2

    # Modifier les eid dans document_eids
    document_eids = [eid.split(".0-")[-1] for eid in document_eids]

    length_list_eids = len(document_eids)
    stop_value = int((length_list_eids-1)/25)

    # Extraire les données du premier élément obligatoirement à part sinon cela impacte la boucle for si length > 25 !
    co = CitationOverview(identifier=document_eids[0:1], start=first_year, end=first_year+total_annees-1, refresh=True)
    header_citation = co._header
    citation_overviews = []
    citation_overviews.append(co.cc)

    # Plusieurs extractions nécessaires si documents > 25
    if length_list_eids > 25:
        # Boucle le nombre de fois où il y a 25 dans length_list_eids
        for i in range(0, int(length_list_eids / 25)):
            # Extraire par 25 les données et les ajouter à la liste principale
            co = CitationOverview(identifier=document_eids[(i*25)+1 : (i*25)+26], start=first_year, end=first_year+total_annees-1, refresh=True)
            header_citation = co._header
            citation_overviews.append(co.cc)

    # Extraire les données qu'il reste (nb de documents < 25)
    if (length_list_eids-1)%25 != 0:
        co = CitationOverview(identifier=document_eids[stop_value*25 + 1 : stop_value*25 + (length_list_eids-1)%25 + 1], start=first_year, end=first_year+total_annees-1, refresh=True)
        header_citation = co._header
        citation_overviews.append(co.cc)
    else:
        citation_overviews.append([[(0, 0) for _ in range(total_annees)]])

    # Décapsulation de citation_overviews (liste principale) pour faire le total par année
    nb_cit_annees = [0] * total_annees
    for i in range(0, stop_value + 2):
        for y in range(0, len(citation_overviews[i])):
            for z in range(0, len(citation_overviews[i][y])):
                nb_cit_annees[z] += citation_overviews[i][y][z][1]

    # Affichage de manière tabulaire le nb de citations par année (T pour transposition matricielle)
    df = pd.DataFrame(nb_cit_annees).T # Transposition
    years_list = [first_year + i for i in range(0, total_annees)] # Modification des index de colonnes avec les années
    df.columns = years_list

    # Renommer l'axe des colonnes transposées ainsi que le nom de la colonne de données
    df = df.rename_axis('Année', axis='columns')
    df = df.rename(index={0: 'Citations'})

    nb_cit_annees.append(sum(nb_cit_annees))

    return nb_cit_annees, years_list, header_citation

# Fonction qui retourne le tableau pour le graphique des citations
def tab_graph_citations(console: QPlainTextEdit, au_retrieval: AuthorRetrieval, eids_list: list, list2: list, window_width: int):
    ### PARTIE sur les citations
    list1, years_list, header = donnees_citations_graph_citations(au_retrieval, eids_list)
    # Créer une liste de paires avec les éléments alignés
    resultat = list(zip(list1, list2))

    # Affichage de manière tabulaire le nb de citations par année (T pour transposition matricielle)
    df = pd.DataFrame(resultat).T # Transposition
    years_list.append('TOTAL')
    df.columns = years_list

    # Renommer l'axe des colonnes transposées ainsi que le nom de la colonne de données
    df = df.rename_axis('Année', axis='columns')
    df = df.rename(index={0: 'Citations'})
    df = df.rename(index={1: 'Documents'})
    
    console.append("\n" + '<p style="text-decoration: underline; color: black;">Tableau pour le graphique des <b>Citations</b>:</p>')
    console.append(df.to_string(index=True, col_space=0, line_width=window_width))

    return df, [au_retrieval.given_name, au_retrieval.surname], header


# Fonction qui retourne les valeurs de l'encadré du rapport en fonction de l'eid du chercheur sélectionné
def valeurs_encadre(console: QPlainTextEdit, author_eid, years_list: list):
    au = AuthorLookup(author_id=author_eid, refresh=True)

    liste_sch_out = _replace_none_with_zero(au.get_metrics_Other(metricType='ScholarlyOutput', yearRange='10yrs').List)
    index_10y_adapted = liste_sch_out[0].index(years_list[0]) if years_list[0] in liste_sch_out[0] else 0

    liste_sch_out_10y_adapted = liste_sch_out[-1][index_10y_adapted:9]
    annee_10y_adapt = liste_sch_out[0][index_10y_adapted]


    tot_scholarly_5y = sum(liste_sch_out[-1][4:9])
    tot_scholarly_10y = sum(liste_sch_out_10y_adapted)
    tot_scholarly_5yc = sum([liste_sch_out[-1][4]] + _replace_none_with_zero(au.get_metrics_Other(metricType='ScholarlyOutput', yearRange='5yrsAndCurrent').List[-1]))


    tot_top_citations = sum(_replace_none_with_zero(au.get_metrics_Percentile(metricType='OutputsInTopCitationPercentiles', yearRange='10yrs').List[-1][4:9]))
    top_citations = round(tot_top_citations/tot_scholarly_5y*100, 1) if tot_scholarly_5y != 0 else 0

    tot_acad_collab = sum(_replace_none_with_zero([au.get_metrics_Collaboration(metricType='AcademicCorporateCollaboration', yearRange='10yrs', collabType='Academic-corporate collaboration').List[-1][4]] + au.get_metrics_Collaboration(metricType='AcademicCorporateCollaboration', yearRange='5yrsAndCurrent', collabType='Academic-corporate collaboration').List[-1]))
    acad_collab = round(tot_acad_collab/tot_scholarly_5yc*100, 1) if tot_scholarly_5yc != 0 else 0


    liste_sch_out_10y_adapted_ArticlesConf = _replace_none_with_zero(au.get_metrics_Other(metricType='ScholarlyOutput', yearRange='10yrs', includedDocs='ArticlesConferencePapers').List)[-1][index_10y_adapted:9]
    tot_liste_sch_out_10y_adapted_ArticlesConf = sum(liste_sch_out_10y_adapted_ArticlesConf)

    liste_cit_per_pub = _replace_none_with_zero(au.get_metrics_Other(metricType='CitationsPerPublication', yearRange='10yrs', includedDocs='ArticlesConferencePapers').List[-1][index_10y_adapted:9])
    cit_per_pub = round(sum([elem1 * elem2 for elem1, elem2 in zip(liste_sch_out_10y_adapted_ArticlesConf, liste_cit_per_pub)])/tot_liste_sch_out_10y_adapted_ArticlesConf, 1) if tot_liste_sch_out_10y_adapted_ArticlesConf != 0 else 0

    liste_moy_MCR = _replace_none_with_zero(au.get_metrics_Other(metricType='FieldWeightedCitationImpact', yearRange='10yrs', includedDocs='ArticlesConferencePapers').List[-1][index_10y_adapted:9])
    moy_MCR = round(sum([elem1 * elem2 for elem1, elem2 in zip(liste_sch_out_10y_adapted_ArticlesConf, liste_moy_MCR)])/tot_liste_sch_out_10y_adapted_ArticlesConf, 2) if tot_liste_sch_out_10y_adapted_ArticlesConf != 0 else 0

    return [top_citations, cit_per_pub, moy_MCR, acad_collab], annee_10y_adapt, au._header


# Fonction qui permet de retourner un booléen pour connaitre la validité de la commande de l'utilisateur,
# une liste de listes des plages d'années sélectionnées et un DataFrame avec les types de docs sélectionnés
def selection_plages_annees(console: QPlainTextEdit, annees_selec: str, years: list, df: pd.DataFrame, selected_types: list):
    # Séparer les types de documents sélectionnés par l'utilisateur (et supprimer les espaces avant et après les éléments)
    parts = annees_selec.split(',')
    parts = [element.strip() for element in parts]

    tout_valide = True

    if len(parts) == 1 and parts[0] == "":
        console.append('')
        console.append('<p><a style="font-weight: bold;">Votre sélection:</a> 3ans ({}), 5ans ({}) et Carrière ({})</p>'.format(years[-1] - 4, years[-1] - 6, years[0]))
        console.append('')
        console.append('')
        console.append('<p style="text-decoration: underline; color: black;">Nb de documents en fonction de leur type:</p>')
        # Filtrer le DataFrame en utilisant la méthode isin() avec la liste des index
        df_filtre = df[df.index.isin(selected_types)]
        console.append(df_filtre.to_string(index=True, col_space=0, line_width=200)) # .to_string().encode('utf-8')
        parts = [years[0], years[-1] - 6, years[-1] - 4]
        year_list = []
        for i in range(len(parts)):
            year_range = list(range(parts[i], datetime.now().year + 2))
            year_range.sort(reverse=True)
            year_list.append(year_range)
        return True, year_list, df_filtre # Liste des années par défaut
    
    if not(len(parts) == 2 or len(parts) == 3):
        # Ici indique erreur: manque ou surplus d'éléments
        console.append("<p style={}>! Manque ou surplus d'éléments (2 ou 3 éléments demandés)</p>".format(text_style_warning))
        return False, parts, df
    
    for year_index in parts:
        # Vérifier si l'index est valide
        if not (year_index.isdigit() and years[0] <= int(year_index) <= years[-1]):
            console.append('<p style={}>! Index non valide: {}</p>'.format(text_style_warning, year_index))
            tout_valide = False
            
    if not tout_valide:
        return False, parts, df

    parts = [int(x) for x in parts]
    if len(parts) == 2:
        parts.append(years[0])
    parts.sort(reverse=True)
    
    console.append('')
    console.append('<p><a style="font-weight: bold;">Votre sélection:</a> {}ans ({}), {}ans ({}) et Carrière ({})</p>'.format(years[-1] - parts[0] - 1, parts[0], years[-1] - parts[1] - 1, parts[1], parts[2]))
    console.append('')
    parts.sort(reverse=False)
    
    year_list = []
    for i in range(len(parts)):
        year_range = list(range(parts[i], datetime.now().year + 2))
        year_range.sort(reverse=True)
        year_list.append(year_range)
    console.append('')
    console.append('<p style="text-decoration: underline; color: black;">Nb de documents en fonction de leur type:</p>')
    # Filtrer le DataFrame en utilisant la méthode isin() avec la liste des index
    df_filtre = df[df.index.isin(selected_types)]
    console.append(df_filtre.to_string(index=True, col_space=0, line_width=200)) # .to_string().encode('utf-8')

    return True, year_list, df_filtre


# Fonction utilitaire pour créer une liste de listes en fonction des combinaisons sélectionnées
def _combine_types(chaine: str):
    main_indices_list = []
    selected_types = chaine.split(',')
    selected_types = [element.strip() for element in selected_types]
    
    for types in selected_types:
        pattern = r'\[(.*?)\]'
        contenu_crochets = re.findall(pattern, types)

        if len(contenu_crochets) > 0:
            element_list = [element_of_element.strip() for element_of_element in contenu_crochets[0].split(';')]

            main_indices_list.append(element_list)
        else:
            main_indices_list.append([types])
    return main_indices_list

# Fonction qui retourne un booléen qui confirme la validité de la commande de l'utilisateur
# et les types de docs sélectionnés pour être mis en avant (combinaisons comprises)
def selection_2_types_docs(console: QPlainTextEdit, df: pd.DataFrame, index_took: str):
    # Combiner des types si c'est demandé
    selected_types = _combine_types(index_took)
    
    tout_valide = True
    liste_types_selec = df['Type de documents'].to_list()

    if len(selected_types) == 1 and selected_types[0][0] == "":
        console.append('')
        console.append('<p><a style="font-weight: bold;">Votre sélection:</a> {}, {}</p>'.format(liste_types_selec[0], liste_types_selec[1] if len(liste_types_selec)>1 else '∅'))
        console.append('')
        return True, [[liste_types_selec[0]], [liste_types_selec[1] if len(liste_types_selec)>1 else '∅']]
    
    flattened_list = [element for sublist in selected_types for element in sublist]
    for type_index in flattened_list:
        # Vérifier si l'index est valide
        if not (type_index.isdigit() and int(type_index) < len(df)):
            console.append('<p style={}>! Index non valide: {}</p>'.format(text_style_warning, type_index))
            tout_valide = False
            continue
        elif flattened_list.count(str(int(type_index))) > 1:
            console.append('<p style={}>! Doublon trouvé: {}</p>'.format(text_style_warning, type_index))
            tout_valide = False

    if not len(selected_types) == 2:
        console.append("<p style={}>! Manque ou surplus d'éléments (2 éléments demandés)</p>".format(text_style_warning))
        return False, selected_types

    if tout_valide:
        # Obtenir une liste d'entier puis mettre à jour la liste
        selected_types = [[int(element) for element in sublist] for sublist in selected_types]

        console.append('')
        console.append('<p><a style="font-weight: bold;">Votre sélection:</a> {}, {}</p>'.format(df.at[selected_types[0][0], 'Type de documents'], df.at[selected_types[1][0], 'Type de documents']))
        console.append('')
        selected_types = [[df.loc[element, 'Type de documents'] for element in sublist] for sublist in selected_types]
    
    return tout_valide, selected_types




# Fonction qui retourne le tableau pour le graphique des publications
def tab_graph_publications(console: QPlainTextEdit, au_retrieval: AuthorRetrieval, document_eids: list, liste_annees: list, liste_type: list, window_width: int):
    # Parcourir chaque liste interne et modifier les années de int à str
    for sous_liste in liste_annees:
        # Parcourir chaque élément de la liste interne et le convertir en str
        for i in range(len(sous_liste)):
            sous_liste[i] = str(sous_liste[i])

    # Filtrer les documents en fonction des EIDs spécifiés
    docs = pd.DataFrame(au_retrieval.get_documents(refresh=10))
    docs_filtered = docs[docs['eid'].isin(document_eids)]

    # Extraire les années de publication
    df = pd.DataFrame({
        'DocType': docs_filtered['subtypeDescription'],
        'Year': pd.to_datetime(docs_filtered['coverDate']).dt.year.astype(str),
    })
    liste_type_en = [[trad_fr2en[doc] for doc in sublist] for sublist in liste_type]

    flattened_liste_type_en = [element for sublist in liste_type_en for element in sublist]
    results = pd.DataFrame(columns=flattened_liste_type_en + ['Autres'])  # DataFrame vide pour stocker les résultats

    for annees in liste_annees:
        df_filtered = df[df['Year'].isin(annees)]
        counts = df_filtered['DocType'].value_counts()
        total_doctype = len(df_filtered)

        liste = counts.reindex(flattened_liste_type_en, fill_value=0).tolist()
        liste.append(total_doctype - sum(liste))

        df_result = pd.DataFrame([liste], columns=flattened_liste_type_en + ['Autres'])
        results = pd.concat([results, df_result], ignore_index=True)

    for sublist in liste_type_en:
        for i in range(1, len(sublist)):
            results[sublist[0]] += results[sublist[i]]
            results = results.drop(sublist[i], axis=1)

    # Renommage dynamique des index et colonnes du dataframe
    results = results.rename(index={0: 'Carrière (≥' + liste_annees[0][-1] + ')', 1: liste_annees[1][-1] + ' à ≥' + liste_annees[0][1], 2: liste_annees[2][-1] + ' à ≥' + liste_annees[0][1]}, 
                             columns={liste_type_en[0][0]: liste_type[0][0], liste_type_en[1][0]: liste_type[1][0]})

    # Calculer le total par ligne et ajouter une colonne 'Total'
    results = results.assign(TOTAL=results.sum(axis=1))

    # Afficher le tableau
    console.append("\n" + '<p style="text-decoration: underline; color: black;">Tableau pour le graphique des <b>Publications</b>:</p>')
    console.append(results.to_string(index=True, col_space=0, line_width=window_width))

    return results

# Fonction utilitaire qui permet d'une liste de retourner une liste avec des 0
# à la place des éléments vides (NONE)
def _replace_none_with_zero(lst: list):
    for i in range(len(lst)):
        if lst[i] is None:
            lst[i] = 0
    return lst


# Fonction utilitaire de la fonction "tab_graph_SNIP" pour permettre d'extraire depuis un résultat
# d'une requête les valeurs nécessaires pour les calculs pour le graphique SNIP
def _for_SNIP_list_10y_current_future(lst: list):
    for element in lst:
        value_by_year = element['valueByYear']

        if element['threshold'] == 5:
            annees = [int(annee) for annee in list(value_by_year.keys())]
            element_with_threshold_5 = list(value_by_year.values())
        elif  element['threshold'] == 10:
            element_with_threshold_10 = list(value_by_year.values())
        elif  element['threshold'] == 25:
            element_with_threshold_25 = list(value_by_year.values())

    return [annees, element_with_threshold_5, element_with_threshold_10, element_with_threshold_25]

# Fonction qui retourne un DataFrame (tableau) pour le graphique SNIP du rapport
def tab_graph_SNIP(console: pd.DataFrame, author_id: str, years_list: list, window_width: int):
    years_list = [[int(item) for item in sublist] for sublist in years_list]

    au = AuthorLookup(author_id=author_id, refresh=True)

    tenycf_list = _for_SNIP_list_10y_current_future(au._get_metrics_rawdata(metricType='PublicationsInTopJournalPercentiles', yearRange='10yrs', journalImpactType="SNIP", includedDocs='ArticlesReviews'))
    fivey_c_f_list = _for_SNIP_list_10y_current_future(au._get_metrics_rawdata(metricType="PublicationsInTopJournalPercentiles", yearRange="3yrsAndCurrentAndFuture", journalImpactType="SNIP", includedDocs='ArticlesReviews'))

    teny_scho_list = au.get_metrics_Other(metricType="ScholarlyOutput", yearRange="10yrs", includedDocs='ArticlesReviews').List
    fivey_c_f_scho_list = au.get_metrics_Other(metricType="ScholarlyOutput", yearRange="3yrsAndCurrentAndFuture", includedDocs='ArticlesReviews').List

    tenycf_list = [_replace_none_with_zero(item1 + item2[-2:]) for item1, item2 in zip(teny_scho_list, fivey_c_f_scho_list)] + [_replace_none_with_zero(item1 + item2[-2:]) for item1, item2 in zip(tenycf_list, fivey_c_f_list)][-3:]


    default_index_year_list = [0, -7, -5]
    
    real_years_list = [years_list[i][-1] if years_list[i][-1] in tenycf_list[0] else tenycf_list[0][default_index_year_list[i]] for i in range(3)]

    column_name_list = ['Top 5%', 'Top 10%', 'Top 25%', 'Autres']
    df = pd.DataFrame().assign(**{column_name_list[i]: [sum(tenycf_list[i+2][tenycf_list[0].index(year):]) for year in real_years_list] for i in range(len(column_name_list) - 1)})

    # Création de la colonne Autres et réalisation des vraies colonnes de données Top 25% et Top 10%
    df[column_name_list[3]] = [sum(tenycf_list[1][tenycf_list[0].index(year):]) - df['Top 25%'].to_list()[j] for j, year in enumerate(real_years_list)]
    df['Top 25%'] = df['Top 25%'] - df['Top 10%']
    df['Top 10%'] = df['Top 10%'] - df['Top 5%']

    # Renommez les index
    df = df.rename(index={0: '≥'+str(real_years_list[0]), 1: str(real_years_list[1])+' à ≥'+str(tenycf_list[0][-2]), 2: str(real_years_list[2])+' à ≥'+str(tenycf_list[0][-2])})

    # Ajouter une colonne "Total" contenant la somme des valeurs des autres colonnes
    df['TOTAL'] = df.sum(axis=1)

    # Afficher le tableau
    console.append('\n')
    console.append('<p style="text-decoration: underline; color: black;">Tableau pour le graphique <b>SNIP</b>:</p>')
    console.append(df.to_string(index=True, col_space=0, line_width=window_width))

    return df, au._header


# Fonction utilitaire de la fonction "tab_graph_Collab" pour permettre d'extraire depuis un résultat
# d'une requête les valeurs nécessaires pour les calculs pour le graphique SNIP
def _for_Collab_list_10y_current_future(lst: list):
    # print(lst)
    for element in lst:
        value_by_year = element['valueByYear']
        if element['collabType'] == "Institutional collaboration":
            annees = [int(annee) for annee in list(value_by_year.keys())]
            inst_collab = list(value_by_year.values())
        elif  element['collabType'] == "International collaboration":
            international_collab = list(value_by_year.values())
        elif  element['collabType'] == "National collaboration":
            national_collab = list(value_by_year.values())
        elif  element['collabType'] == "Single authorship":
            no_collab = list(value_by_year.values())

    return [annees, inst_collab, international_collab, national_collab, no_collab]

# Fonction qui retourne un DataFrame (tableau) pour le graphique Collaborations du rapport
def tab_graph_Collab(console: pd.DataFrame, author_id: str, years_list: list, window_width: int):
    au = AuthorLookup(author_id=author_id, refresh=True)

    tenycf_list = _for_Collab_list_10y_current_future(au._get_metrics_rawdata(metricType='Collaboration', yearRange='10yrs'))
    fivey_c_f_list = _for_Collab_list_10y_current_future(au._get_metrics_rawdata(metricType='Collaboration', yearRange='3yrsAndCurrentAndFuture'))

    tenycf_list = [_replace_none_with_zero(item1 + item2[-2:]) for item1, item2 in zip(tenycf_list, fivey_c_f_list)]

    default_index_year_list = [0, -7, -5]
    
    real_years_list = [years_list[i][-1] if years_list[i][-1] in tenycf_list[0] else tenycf_list[0][default_index_year_list[i]] for i in range(3)]

    column_name_list = ['Internat.', 'Nationale', 'Inst.', 'Aucune']
    df = pd.DataFrame().assign(**{column_name_list[i]: [sum(tenycf_list[i+1][tenycf_list[0].index(year):]) for year in real_years_list] for i in range(len(column_name_list))})

    # Intervertir les colonnes pour avoir l'ordre: International, National, Institutionnel
    df['Internat.'], df['Nationale'], df['Inst.'] = df['Nationale'].copy(), df['Inst.'].copy(), df['Internat.'].copy()

    # Renommez les index
    df = df.rename(index={0: '≥'+str(real_years_list[0]), 1: str(real_years_list[1])+' à ≥'+str(tenycf_list[0][-2]), 2: str(real_years_list[2])+' à ≥'+str(tenycf_list[0][-2])})

    # Ajouter une colonne "Total" contenant la somme des valeurs des autres colonnes
    df['TOTAL'] = df.sum(axis=1)

    # Afficher le tableau
    console.append('\n')
    console.append('<p style="text-decoration: underline; color: black;">Tableau pour le graphique des <b>Collaborations</b>:</p>')
    console.append(df.to_string(index=True, col_space=0, line_width=window_width))

    return df, au._header




# Fonction qui permet d'exporter les données sur le gabarit Excel et d'appeler les
# routines VBA du gabarit
def Excel_part1(df: pd.DataFrame, nom_prenom: list, en_tete: list, annee_10y_adapt: int):
    # Ouvrir le classeur Excel existant
    nom_fichier = os.path.dirname(os.path.abspath(__file__)) + '\\..\\GABARIT.xlsm'
    nom_feuille = 'Raw_Data'
    nom_module = 'Module1'
    nom_procedure = 'AjusterDynamiquementAbscisseGraphiqueCitations'

    cell_tab_citations = [4, 0] # ligne, colonne

    # Création de l'objet Excel, et le rendre visible en plein écran lors du processus
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.WindowState = win32.constants.xlMaximized

    try:
        # Vérifier si le fichier Excel est déjà ouvert
        for wb in excel.Workbooks:
            if wb.FullName == nom_fichier:
                wb.Close(False)  # Fermer le classeur sans enregistrer les modifications

        # Ouverture du fichier Excel
        classeur = excel.Workbooks.Open(nom_fichier)
        classeur.Visible = True  # Rendre le classeur visible
        classeur.WindowState = win32.constants.xlMaximized  # Mettre le classeur en plein écran 

        # Enregistrer sous le nouveau nom
        date_formated = datetime.now().strftime('%Y-%m-%d')
        # print(date_formated)
        nom_ou_prenom = []
        for i in range(len(nom_prenom)):
            # Supprimer tous les accents, ...
            nom_ou_prenom.append(unicodedata.normalize("NFD", nom_prenom[i]).encode("ascii", "ignore").decode("utf-8")) # une str sans accent
            # Remplacer les espaces par des tirets
            nom_ou_prenom[i] = nom_ou_prenom[i].replace(' ', '-')
            # Supprimer tous les apostrophes
            nom_ou_prenom[i] = nom_ou_prenom[i].replace("'", '')

        classeur.SaveAs(os.path.abspath(DOCS_PATH[0] + '/' + date_formated + '_' + nom_ou_prenom[1] + '_' + nom_ou_prenom[0] + '.xlsm'), FileFormat=52)

        # Mettre la fenêtre en premier plan
        try:
            win32gui.SetForegroundWindow(win32gui.FindWindow(None, classeur.Name + " - Excel"))
        except:
            win32gui.SetForegroundWindow(win32gui.FindWindow(None, classeur.Name.split('.')[0] + " - Excel"))
      
        # Accéder à la feuille de calcul existante
        feuille = classeur.Worksheets(nom_feuille)

        # Effacer le contenu des lignes de cellules et cellules à modifier
        start_row = 2 + cell_tab_citations[0]  # Numéro de la première ligne de cellule (à supprimer)
        end_row = start_row + len(df) - 1  # Numéro de la dernière ligne de cellule (à supprimer)

        feuille.Range(f"{start_row-1}:{end_row}").ClearContents()  # Nettoie seulement le contenu des lignes souhaitées

        # Path ainsi que le path de ce programme pour enregistrer le Word
        feuille.Cells(1, 110).Value = DOCS_PATH[0]
        feuille.Cells(2, 110).Value = os.path.dirname(os.path.abspath(__file__)) + "\\..\\"

        # Écrire prénom et nom
        for i in range(len(nom_prenom)):
            feuille.Cells(1 + i, 2).Value = nom_prenom[i]

        # Écrire année des 10y_adapt
        feuille.Cells(1, 104).Value = annee_10y_adapt

        # Écrire prénom et nom sans accent
        for i in range(len(nom_prenom)):
            feuille.Cells(1 + i, 5).Value = nom_ou_prenom[i]

        # Écrire les données du DataFrame dans la feuille de calcul
        for i, row in enumerate(df.values):
            for j, value in enumerate(row):
                if j > len(row) - 4:
                    feuille.Cells(i+2+cell_tab_citations[0], j+2+cell_tab_citations[1]+2).Value = value  # Ajouter 2 pour décaler les cellules
                    if j > len(row) - 3:
                        continue
                feuille.Cells(i+2+cell_tab_citations[0], j+2+cell_tab_citations[1]).Value = value

        # Écrire les index dans la première colonne de la feuille de calcul
        for i, index_value in enumerate(df.index):
            feuille.Cells(i+2+cell_tab_citations[0], 1+cell_tab_citations[1]).Value = index_value

        # Écrire les données supplémentaires
        year_list = df.columns.to_list()
        feuille.Cells(1, 103).Value = year_list[-3]
        feuille.Cells(2, 102).Value = year_list[0]
        feuille.Cells(1, 106).Value = date_formated

        # Écrire les noms des colonnes dans la première ligne de la feuille de calcul
        if len(year_list) - 2 <= 20:
            for j, column_name in enumerate(df.columns):
                if j == len(df.columns) - 1:
                    feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1] + 2).Value = column_name
                    continue
                elif j > len(row) - 4:
                    feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1] + 2).Value = str(column_name)[-2:]
                    if j > len(row) - 3:
                        continue
                feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1]).Value = str(column_name)[-2:]
            if len(year_list) - 2 > 10:
                for j, column_name in enumerate(df.columns):
                    if j % 2 == 1 and j != len(year_list) - 3 or (len(year_list) - 2) % 2 == 0 and j == len(year_list) - 4:
                        feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1]).Value = None
        
        elif len(year_list) - 2 < 30:
            for j, column_name in enumerate(df.columns):
                if j > len(row) - 4:
                    feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1] + 2).Value = column_name
                    if j > len(row) - 3:
                        continue
                feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1]).Value = column_name
            for j, column_name in enumerate(df.columns):
                if j % 2 == 1 and j != len(year_list) - 3 or (len(year_list) - 2) % 2 == 0 and j == len(year_list) - 4:
                    feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1]).Value = None
        
        else:
            for j, column_name in enumerate(df.columns):
                if j > len(row) - 4:
                    feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1] + 2).Value = column_name
                    if j > len(row) - 3:
                        continue
                feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1]).Value = column_name
            for j, column_name in enumerate(df.columns):
                if j % 5 != 0 and j != len(year_list) - 3 or (len(year_list) - 2) % 2 == 0 and j == len(year_list) - 5 or len(year_list) - 7 < j < len(year_list) - 3:
                    feuille.Cells(1 + cell_tab_citations[0], j + 2 + cell_tab_citations[1]).Value = None

        # Écrire les données de l'en-tête de SciVal
        for i in range(len(en_tete)):
            feuille.Cells(32 + i, 2).Value = en_tete[i]

        # Appel de la procédure VBA
        excel.Run(f'{nom_module}.{nom_procedure}', nom_feuille, 'TOTAL', 5)

        # Attend confirmation du user
        while feuille.Cells(1, 101).Value != "OK":
            pass
        
        classeur.Visible = False  # Rendre le classeur visible
        excel.Visible = False

    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

    return excel, classeur

    
# Fonction qui reprend le classeur ouvert (caché) et qui permet d'exporter le reste des 
# données sur le gabarit Excel et d'appeler la routine VBA du gabarit Excel qui
# remplie le gabarit Word pour avoir la fiche bibliométrique finale !
def Excel_part2(excel, classeur, df: pd.DataFrame, df_SNIP: pd.DataFrame, df_Collab: pd.DataFrame):
    # Ouvrir le classeur Excel existant
    nom_fichier = os.path.dirname(os.path.abspath(__file__)) + '\\..\\GABARIT.xlsm'
    nom_feuille = 'Raw_Data'
    nom_module = 'Module1'

    cell_tab_publications = [10, 41, 48] # lignes de commencement des tableaux pour les graphiques

    # Rerendre visible l'Excel
    excel.Visible = True

    try:
        # Vérifier si le fichier Excel est déjà ouvert
        for wb in excel.Workbooks:
            if wb.FullName == nom_fichier:
                wb.Close(False)  # Fermer le classeur sans enregistrer les modifications

        # Ouverture du fichier Excel
        classeur.Visible = True  # Rendre le classeur visible
        classeur.WindowState = win32.constants.xlMaximized  # Mettre le classeur en plein écran 

        # Mettre la fenêtre Excel en premier plan
        try:
            win32gui.SetForegroundWindow(win32gui.FindWindow(None, classeur.Name + " - Excel"))
        except:
            win32gui.SetForegroundWindow(win32gui.FindWindow(None, classeur.Name.split('.')[0] + " - Excel"))


        # Accéder à la feuille de calcul existante
        feuille = classeur.Worksheets(nom_feuille)


        ### GRAPHIQUE DES PUBLICATIONS ###

        # Effacer le contenu des lignes de cellules et cellules à modifier
        start_row = 2 + cell_tab_publications[0]  # Numéro de la première ligne de cellule (à supprimer)
        end_row = start_row + len(df) - 1  # Numéro de la dernière ligne de cellule (à supprimer)

        feuille.Range(f"{start_row-1}:{end_row}").ClearContents()   # Nettoie seulement le contenu des lignes souhaitées

        # Écrire les données du DataFrame dans la feuille de calcul
        for i, row in enumerate(df.values):
            for j, value in enumerate(row):
                if not value == 0:
                    feuille.Cells(i+2+cell_tab_publications[0], j+2).Value = value
                else:
                    feuille.Cells(i+2+cell_tab_publications[0], j+2).ClearContents()

        # Écrire les index dans la première colonne de la feuille de calcul
        for i, index_value in enumerate(df.index):
            feuille.Cells(i+2+cell_tab_publications[0], 1).Value = index_value

        # Écrire les noms des colonnes dans la première ligne de la feuille de calcul
        for j, column_name in enumerate(df.columns):
            feuille.Cells(1+cell_tab_publications[0], j+2).Value = column_name


        ### GRAPHIQUE SNIP ###

        # Effacer le contenu des lignes de cellules et cellules à modifier
        start_row = 2 + cell_tab_publications[1]  # Numéro de la première ligne de cellule (à supprimer)
        end_row = start_row + len(df_SNIP) - 1  # Numéro de la dernière ligne de cellule (à supprimer)

        feuille.Range(f"{start_row-1}:{end_row}").ClearContents()   # Nettoie seulement le contenu des lignes souhaitées

        # Écrire les données du DataFrame dans la feuille de calcul
        for i, row in enumerate(df_SNIP.values):
            for j, value in enumerate(row):
                if not value == 0:
                    feuille.Cells(i+2+cell_tab_publications[1], j+2).Value = value
                else:
                    feuille.Cells(i+2+cell_tab_publications[1], j+2).ClearContents()

        # Écrire les index dans la première colonne de la feuille de calcul
        for i, index_value in enumerate(df_SNIP.index):
            feuille.Cells(i+2+cell_tab_publications[1], 1).Value = index_value

        # Écrire les noms des colonnes dans la première ligne de la feuille de calcul
        for j, column_name in enumerate(df_SNIP.columns):
            feuille.Cells(1+cell_tab_publications[1], j+2).Value = column_name


        ### GRAPHIQUE COLLAB ###

        # Effacer le contenu des lignes de cellules et cellules à modifier
        start_row = 2 + cell_tab_publications[2]  # Numéro de la première ligne de cellule (à supprimer)
        end_row = start_row + len(df_Collab) - 1  # Numéro de la dernière ligne de cellule (à supprimer)

        feuille.Range(f"{start_row-1}:{end_row}").ClearContents()   # Nettoie seulement le contenu des lignes souhaitées

        # Écrire les données du DataFrame dans la feuille de calcul
        for i, row in enumerate(df_Collab.values):
            for j, value in enumerate(row):
                if not value == 0:
                    feuille.Cells(i+2+cell_tab_publications[2], j+2).Value = value
                else:
                    feuille.Cells(i+2+cell_tab_publications[2], j+2).ClearContents()

        # Écrire les index dans la première colonne de la feuille de calcul
        for i, index_value in enumerate(df_Collab.index):
            feuille.Cells(i+2+cell_tab_publications[2], 1).Value = index_value

        # Écrire les noms des colonnes dans la première ligne de la feuille de calcul
        for j, column_name in enumerate(df_Collab.columns):
            feuille.Cells(1+cell_tab_publications[2], j+2).Value = column_name

        # Ouvrir Word
        word_app = win32.Dispatch("Word.Application")

        # Fermer les documents
        for doc in word_app.Documents:
            doc.Close(SaveChanges=True)

        # Reboot pour prendre la main
        word_app.Quit()

        # Appel de la procédure VBA
        excel.Run(f'{nom_module}.GenerationWord')

        # Attendre la fin des subroutines en cours
        time.sleep(1)
        
        # Récupérer la feuille à supprimer
        sheet = wb.Sheets("Main")
        # Supprimer la feuille
        excel.DisplayAlerts = False  # Désactiver les alertes
        sheet.Delete()
        excel.DisplayAlerts = True  # Réactiver les alertes

        # Supprimer les cellules tampons
        start_row, num_rows = 1, 2
        start_col, num_cols = 100, 11

        for row in range(start_row, start_row + num_rows):
            for col in range(start_col, start_col + num_cols):
                feuille.Cells(row, col).ClearContents()

        # Enregistrer et fermer le classeur Excel
        classeur.Close(SaveChanges=True)

        # Fermeture de l'application Excel
        excel.Quit()

    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

