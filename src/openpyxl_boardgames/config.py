"""Constantes du projet : chemins, références Excel, formules et configuration des données """

from pathlib import Path

# -------------------------------------------------------------------------------------------
# URL des tables de données
# -------------------------------------------------------------------------------------------

URL1 = "https://minio.lab.sspcloud.fr/xifar19/boardgames/bgg_dataset.csv"
URL2 = "https://minio.lab.sspcloud.fr/xifar19/boardgames/bgg_pictures.csv"
URL3 = "https://minio.lab.sspcloud.fr/xifar19/boardgames/boardgames_ranks.csv"

# -------------------------------------------------------------------------------------------
# Colonnes à sélectionner parmis les données brutes et colonnes à noramliser
# -------------------------------------------------------------------------------------------

COLUMNS = [
    "ID",
    "Name",
    "Year Published",
    "Min Players",
    "Max Players",
    "Play Time",
    "Min Age",
    "Users Rated",
    "Rating Average",
    "BGG Rank",
    "Complexity Average",
    "Owned Users",
    "Mechanics",
    "Domains",
    "artist",
    "designer",
    "publisher",
    "is_expansion",
    "abstracts_rank",
    "cgs_rank",
    "childrensgames_rank",
    "familygames_rank",
    "partygames_rank",
    "strategygames_rank",
    "thematic_rank",
    "wargames_rank",
]

COLUMNS_TO_NORMALIZE = [
    "Rating Average", 
    "Complexity Average", 
    "Users Rated", 
    "Owned Users",
    "Min Players", 
    "Max Players", 
    "Play Time", 
    "Min Age"
]

# -------------------------------------------------------------------------------------------
# Colonne à nettoyer et type à changer
# -------------------------------------------------------------------------------------------

OBJECT_FLOAT = [
    'Complexity Average',
    'Rating Average'
]

FILL_NA = {
    'is_expansion':0,
    'abstracts_rank':0,
    'cgs_rank':0,
    'childrensgames_rank':0,
    'familygames_rank':0,
    'partygames_rank':0,
    'strategygames_rank':0,
    'thematic_rank':0,
    'wargames_rank':0,
    'Owned Users':0,
    'Domains': 'Without domain'
}

NEW_TYPE = {
    'ID':'int64',
    'Year Published':'int64',
    'Rating Average': 'float64', 
    'Complexity Average': 'float64',
    'is_expansion':'int64',
    'abstracts_rank':'int64',
    'cgs_rank':'int64',
    'childrensgames_rank':'int64',
    'familygames_rank':'int64',
    'partygames_rank':'int64',
    'strategygames_rank':'int64',
    'thematic_rank':'int64',
    'wargames_rank':'int64',
    'Owned Users':'int64'
}

# -------------------------------------------------------------------------------------------
# Chemin de l'Excel intermédiaire et de l'Excel final
# -------------------------------------------------------------------------------------------

PATH_TEMPLATE = Path("output/template.xlsx")
OUTPUT_PATH = Path("results/reporting.xlsx")

# -------------------------------------------------------------------------------------------
# Nom des feuilles dans le reporting
# -------------------------------------------------------------------------------------------

SHEET_DATA = "DATA"
SHEET_TDB = "TDB"
SHEET_TABLEAU = "TABLEAU"

# -------------------------------------------------------------------------------------------
# Référence et formule pour la feuille TDB
# -------------------------------------------------------------------------------------------

TITLE_TDB = [
    "Filtre famille",
    "Nombre de jeux",
    "Proportion",
    "Rang BGG médian",
    "Année moyenne de parution",
]

TITLE_CELLS_TDB = [
    "B5",
    "E5",
    "H5",
    "K5",
    "N5"
]

TITLE_MERGE_TDB= [
    "B5:C7",
    "E5:F7",
    "H5:I7",
    "K5:L7",
    "N5:O7"
]

KPI_FORMULA_TDB = [
    '=COUNT(AG:AG)',
    "=E8/20327",
    "=_xlfn.AGGREGATE(12,6,AJ:AJ)",
    "=_xlfn.AGGREGATE(1,6,AC:AC)"
]

KPI_CELLS_TDB = [
    "E8",
    "H8",
    "K8",
    "N8"
]

KPI_MERGE_TDB = [
    "E8:F9",
    "H8:I9",
    "K8:L9",
    "N8:O9"
]

FORMULA_FILTER = '=_xlfn.FILTER(DATA!A2:AH20328,DATA!N2:N20328=TDB!B8,'')'

# -------------------------------------------------------------------------------------------
# Référence et formule pour la feuille TABLEAU
# -------------------------------------------------------------------------------------------

TITLE_TABLEAU = [
    "Jeu",
    "Année de publication",
    "Famille",
    "Rang BGG",
    "Complexité",
    "Min joueur",
    "Max joueur",
    "Temps de jeu",
    "Game designer",
    "Artiste",
    "Editeur",
    "Note normalisée",
    "Complexité normalisée",
    "Min joueur normalisé",
    "Max joueur normalisé",
    "Temps de partie normalisé",
    "Age min normalisé"
]

TITLE_CELLS_TABLEAU =[
    "A1", "B1", "C1", "D1", "E1", "F1", "G1", 
    "H1", "I1", "J1", "K1", "L1", "M1", "N1", 
    "O1", "P1", "Q1"
]

KPI_FORMULA_TABLEAU =[
    "=_xlfn.UNIQUE(TDB!AB:AB)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AC:AC)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AN:AN)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AJ:AJ)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AK:AK)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AD:AD)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AE:AE)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AF:AF)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AP:AP)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AO:AO)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!AQ:AQ)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!BA:BA)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!BB:BB)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!BE:BE)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!BF:BF)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!BG:BG)",
    "=_xlfn.XLOOKUP(A2:A10143,TDB!AB:AB,TDB!BH:BH)"
]

KPI_FORMULA_RANGE =[
    "A2:A10143",
    "B2:B10143",
    "C2:C10143",
    "D2:D10143",
    "E2:E10143",
    "F2:F10143",
    "G2:G10143",
    "H2:H10143",
    "I2:I10143",
    "J2:J10143",
    "K2:K10143",
    "L2:L10143",
    "M2:M10143",
    "N2:N10143",
    "O2:O10143",
    "P2:P10143",
    "Q2:Q10143",
]

KPI_CELLS_TABLEAU =[
    "A2", "B2", "C2", "D2", "E2", "F2", "G2", 
    "H2", "I2", "J2", "K2", "L2", "M2", "N2", 
    "O2", "P2", "Q2"
]

# -------------------------------------------------------------------------------------------
# Référence et formule pour le radar chart
# -------------------------------------------------------------------------------------------

RADCHART_CELLS = [
    "2","3","4","5","6","7"
]

RADCHART_TITLES = [
    "Note moyenne normalisée",
    "Complexité moyenne normalisée",
    "Nombre moyen minimal de joueurs normalisé",
    "Nombre moyen maximal de joueurs normalisé",
    "Temps de partie moyen normalisé",
    "Age minimal moyen normalisé"
]

RADCHART_FORMULA1=[
    '=_xlfn.AGGREGATE(1,6,BA:BA)',
    '=_xlfn.AGGREGATE(1,6,BB:BB)',
    '=_xlfn.AGGREGATE(1,6,BE:BE)',
    '=_xlfn.AGGREGATE(1,6,BF:BF)',
    '=_xlfn.AGGREGATE(1,6,BG:BG)',
    '=_xlfn.AGGREGATE(1,6,BH:BH)'
]

RADCHART_FORMULA2=[
    '=TABLEAU!L2',
    '=TABLEAU!M2',
    '=TABLEAU!N2',
    '=TABLEAU!O2',
    '=TABLEAU!P2',
    '=TABLEAU!Q2'
]