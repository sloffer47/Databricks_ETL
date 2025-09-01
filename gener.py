import pandas as pd
from datetime import datetime

# Fonction pour uniformiser la longueur des colonnes
def normalize_length(data_dict):
    max_len = max(len(v) for v in data_dict.values())
    for k in data_dict:
        data_dict[k] += [""] * (max_len - len(data_dict[k]))
    return data_dict

# Modèle pour le planning journalier et prévisionnel
template_columns = ["HEURES", "TACHES", "COMPREHENSION", ""]
template_hours = [
    "8:00 - 9:00", "9:00 - 10:00", "10:15 - 11:00", "11:00 - 12:00",
    "", "13:00 - 14:00", "14:00 - 15:00", "15:00 - 16:00", "16:15 - 17:15",
    "17:15 - 18:00", "", "18:00 - 19:00", "19:00 - 20:00"
]

# Données pour le compte rendu du 12/08/2025 (Talend) et prévisionnel (exos Talend)
planning_data = {
    "date": "2025-08-12",
    "objectif_journalier": "Configurer Talend pour des exercices intermédiaires",
    "taches_journalier": [
        "Configurer un job pour l'importation de données",
        "Utiliser tMap pour des transformations de données",
        "Gérer des erreurs avec tLogRow",
        "Créer un flux de rendu pour l'export des données",
        "Tester le job avec des données d'exemple",
        "Documenter le processus dans Talend",
        "Sauvegarder le projet dans Git",
        "", "", "", "", "", "", "", ""
    ],
    "elements_maitrises_journalier": [
        "Configuration de jobs Talend",
        "Utilisation de tMap",
        "Gestion des erreurs",
        "Exportation des données",
        "Documentation dans Talend",
        "Sauvegarde dans Git",
        "", "", "", "", "", "", "", ""
    ],
    "elements_non_maitrises_journalier": [
        "", "", "", "", "", "", "", "", "", "", "",
        "Création de composants personnalisés",
        "Gestion des flux complexes",
        "", "", "", ""
    ],
    "objectif_previsionnel": "Travailler sur des exercices intermédiaires sur Talend",
    "taches_previsionnel": [
        "Revoir les exercices intermédiaires fournis",
        "Créer des jobs pour des flux complexes",
        "Appliquer des transformations avancées",
        "Utiliser des conditions dans les flux",
        "Tester l'intégration avec des bases de données externes",
        "Documenter toutes les étapes",
        "", "", "", ""
    ],
    "elements_maitrises_previsionnel": [
        "", "", "", "", "", "", "", "", "", "", "", "", ""
    ],
    "elements_non_maitrises_previsionnel": [
        "Flux complexes",
        "Transformation avancée",
        "Conditions dans les jobs",
        "Intégration avec des sources multiples",
        "", "", "", "", "", "", "", ""
    ]
}

# Générer les fichiers Excel
date = planning_data["date"]
date_obj = datetime.strptime(date, "%Y-%m-%d")
excel_date = (date_obj - datetime(1899, 12, 30)).days

# Planning journalier
journalier_data = {
    "HEURES": ["", "", "", "", "Objectif de la journée :", "", "", "", ""] + template_hours + ["", "Bilan de la journée:"],
    "TACHES": ["PLANNING JOURNALIER", "", "", str(excel_date), planning_data["objectif_journalier"]] + [""] * 4 + planning_data["taches_journalier"],
    "COMPREHENSION": ["Nom : MBANDOU Yorick", "", "", "", "", "", "", "Elements maîtrisés", ""] + planning_data["elements_maitrises_journalier"],
    "": ["", "", "", "", "", "", "", "Elements à revoir", ""] + planning_data["elements_non_maitrises_journalier"]
}
journalier_data = normalize_length(journalier_data)
df_journalier = pd.DataFrame(journalier_data)
df_journalier.to_excel(f"C:/Users/myori/Downloads/Planning étudiant journalier {date}.xlsx", index=False, header=False)

# Planning prévisionnel
previsionnel_data = {
    "HEURES": ["", "", "", "", "Objectif de la journée :", "", "", "", ""] + template_hours + ["", "Bilan de la journée:"],
    "TACHES": ["PLANNING PRÉVISIONNEL", "", "", str(excel_date), planning_data["objectif_previsionnel"]] + [""] * 4 + planning_data["taches_previsionnel"],
    "COMPREHENSION": ["Nom : ", "", "", "", "", "", "", "Elements maîtrisés", ""] + planning_data["elements_maitrises_previsionnel"],
    "": ["", "", "", "", "", "", "", "Elements à revoir", ""] + planning_data["elements_non_maitrises_previsionnel"]
}
previsionnel_data = normalize_length(previsionnel_data)
df_previsionnel = pd.DataFrame(previsionnel_data)
df_previsionnel.to_excel(f"C:/Users/myori/Downloads/Planning étudiant prévisionnel {date}.xlsx", index=False, header=False)

print("Fichiers Excel générés dans C:/Users/myori/Downloads pour la date 12/08/2025.")


# import pandas as pd
# from datetime import datetime

# # Fonction pour uniformiser la longueur des colonnes
# def normalize_length(data_dict):
#     max_len = max(len(v) for v in data_dict.values())
#     for k in data_dict:
#         data_dict[k] += [""] * (max_len - len(data_dict[k]))
#     return data_dict

# # Modèle pour le planning journalier et prévisionnel
# template_columns = ["HEURES", "TACHES", "COMPREHENSION", ""]
# template_hours = [
#     "8:00 - 9:00", "9:00 - 10:00", "10:15 - 11:00", "11:00 - 12:00",
#     "", "13:00 - 14:00", "14:00 - 15:00", "15:00 - 16:00", "16:15 - 17:15",
#     "17:15 - 18:00", "", "18:00 - 19:00", "19:00 - 20:00"
# ]

# # Données pour le compte rendu du 11/08/2025 (Talend) et prévisionnel (exos Talend)
# planning_data = {
#     "date": "2025-08-11",
#     "objectif_journalier": "Configurer Talend avec MySQL et explorer les flux de données",
#     "taches_journalier": [
#         "Configurer la connexion JDBC MySQL dans Talend",
#         "Télécharger et organiser le driver JDBC MySQL dans .m2",
#         "Tester la connexion avec un programme Java simple",
#         "Résoudre l'erreur de timezone (serverTimezone=Europe/Paris)",
#         "Importer les schémas des tables (products, customers, orders)",
#         "Créer un job pour extraction + filtrage (tMysqlInput, tFilterRow)",
#         "Utiliser tMap pour transformations (mapping, calculs)",
#         "Ajouter dé-duplication avec tUniqRow",
#         "Effectuer agrégation avec tAggregateRow (total par produit)",
#         "Trier les résultats avec tSortRow",
#         "Exporter vers Excel avec tFileOutputExcel",
#         "Sauvegarder le job dans Git",
#         "", "", "", "", ""
#     ],
#     "elements_maitrises_journalier": [
#         "Configuration JDBC dans Talend",
#         "Organisation du driver dans .m2",
#         "Test de connexion Java",
#         "Résolution d'erreur timezone",
#         "Importation de schémas",
#         "Extraction et filtrage avec tMysqlInput/tFilterRow",
#         "Transformations avec tMap",
#         "Dé-duplication avec tUniqRow",
#         "Agrégation avec tAggregateRow",
#         "Tri avec tSortRow",
#         "Export vers Excel",
#         "Sauvegarde dans Git",
#         "", "", "", "", ""
#     ],
#     "elements_non_maitrises_journalier": [
#         "", "", "", "", "", "", "", "", "", "", "",
#         "Création de composants personnalisés",
#         "Gestion de flux complexes",
#         "", "", "", ""
#     ],
#     "objectif_previsionnel": "Réaliser des exercices supplémentaires sur Talend",
#     "taches_previsionnel": [
#         "Revoir les exos demandés au formateur",
#         "Créer un job pour un flux de données plus complexe",
#         "Appliquer des transformations avancées avec tMap",
#         "Utiliser tAggregateRow pour des calculs personnalisés",
#         "Ajouter des conditions avec tFilterRow",
#         "Tester l'intégration avec d'autres bases ou fichiers",
#         "Résoudre des erreurs potentielles",
#         "Documenter les exos réalisés",
#         "Explorer la création de composants personnalisés en Java",
#         "", "", "", ""
#     ],
#     "elements_maitrises_previsionnel": [
#         "", "", "", "", "", "", "", "", "", "", "", "", ""
#     ],
#     "elements_non_maitrises_previsionnel": [
#         "Exos avancés",
#         "Flux complexes",
#         "Transformations personnalisées",
#         "Conditions avancées",
#         "Intégration multi-sources",
#         "", "", "", "", "", "", "", ""
#     ]
# }

# # Générer les fichiers Excel
# date = planning_data["date"]
# date_obj = datetime.strptime(date, "%Y-%m-%d")
# excel_date = (date_obj - datetime(1899, 12, 30)).days

# # Planning journalier
# journalier_data = {
#     "HEURES": ["", "", "", "", "Objectif de la journée :", "", "", "", ""] + template_hours + ["", "Bilan de la journée:"],
#     "TACHES": ["PLANNING JOURNALIER", "", "", str(excel_date), planning_data["objectif_journalier"]] + [""] * 4 + planning_data["taches_journalier"],
#     "COMPREHENSION": ["Nom : MBANDOU Yorick", "", "", "", "", "", "", "Elements maîtrisés", ""] + planning_data["elements_maitrises_journalier"],
#     "": ["", "", "", "", "", "", "", "Elements à revoir", ""] + planning_data["elements_non_maitrises_journalier"]
# }
# journalier_data = normalize_length(journalier_data)
# df_journalier = pd.DataFrame(journalier_data)
# df_journalier.to_excel(f"C:/Users/myori/Downloads/Plannin etudiant journalier {date}.xlsx", index=False, header=False)

# # Planning prévisionnel
# previsionnel_data = {
#     "HEURES": ["", "", "", "", "Objectif de la journée :", "", "", "", ""] + template_hours + ["", "Bilan de la journée:"],
#     "TACHES": ["PLANNING JOURNALIER", "", "", str(excel_date), planning_data["objectif_previsionnel"]] + [""] * 4 + planning_data["taches_previsionnel"],
#     "COMPREHENSION": ["Nom : ", "", "", "", "", "", "", "Elements maîtrisés", ""] + planning_data["elements_maitrises_previsionnel"],
#     "": ["", "", "", "", "", "", "", "Elements à revoir", ""] + planning_data["elements_non_maitrises_previsionnel"]
# }
# previsionnel_data = normalize_length(previsionnel_data)
# df_previsionnel = pd.DataFrame(previsionnel_data)
# df_previsionnel.to_excel(f"C:/Users/myori/Downloads/Plannin etudiant previsionnelle {date}.xlsx", index=False, header=False)

# print("Fichiers Excel générés dans C:/Users/myori/Downloads pour la date 11/08/2025.")