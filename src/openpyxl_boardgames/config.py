"""Configuration file for the project"""

URL1 = "https://minio.lab.sspcloud.fr/xifar19/boardgames/bgg_dataset.csv"

URL2 = "https://minio.lab.sspcloud.fr/xifar19/boardgames/bgg_pictures.csv"

URL3 = "https://minio.lab.sspcloud.fr/xifar19/boardgames/boardgames_ranks.csv"

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

OBJECT_FLOAT = ['Complexity Average','Rating Average']

FILL_NA = {'is_expansion':0,
        'abstracts_rank':0,
        'cgs_rank':0,
        'childrensgames_rank':0,
        'familygames_rank':0,
        'partygames_rank':0,
        'strategygames_rank':0,
        'thematic_rank':0,
        'wargames_rank':0,
        'Owned Users':0,
        'Domains': 'Without domain'}

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

HIDDEN = [
    'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 
    'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 
    'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH'
]

PATH_TEMPLATE = "output/template.xlsx"

SHEET_DATA = "DATA"
SHEET_TDB = "TDB"

TITLE_TDB = [
    "Filtre famille",
    "Nombre de jeux",
    "Proportion",
    "Rang BGG médian",
    "Année moyenne de parution",
]

TITLE_CELLS_TDB = ["B5","E5","H5","K5","N5"]

TITLE_MERGE_TDB= ["B5:C7","E5:F7","H5:I7","K5:L7","N5:O7"]

KPI_FORMULA = [
    '=10143-COUNTIF(AG:AG,"<>"&NA())',
    "=E8/23327",
    "=_xlfn.AGGREGATE(12,6,AJ:AJ)",
    "=_xlfn.AGGREGATE(1,6,AC:AC)"
]

KPI_CELLS_TDB = ["E8","H8","K8","N8"]

KPI_MERGE_TDB = ["E8:F9","H8:I9","K8:L9","N8:O9"]

FORMULA_FILTER = '=_xlfn.FILTER(DATA!A:AH,DATA!N:N=TDB!B8,"")'

MAX_ROW='10143'