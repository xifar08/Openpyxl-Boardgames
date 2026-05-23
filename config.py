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
        'Owned Users':0}

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