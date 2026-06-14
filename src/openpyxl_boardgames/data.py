"""Chargement et nettoyage des données source."""

from openpyxl_boardgames.components.utils import _clean_data, normalize
from openpyxl_boardgames.config import (
    URL1, 
    URL2, 
    URL3, 
    COLUMNS, 
    OBJECT_FLOAT, 
    FILL_NA, 
    NEW_TYPE,
    COLUMNS_TO_NORMALIZE)
import pandas as pd

def load_data() -> pd.DataFrame:
    """Charge les csv depuis leur URL respectives définies dans la config 
    puis merge en un seul dataframe

    Returns:
        pd.DataFrame: Dataframe brut
    """
    df_main = pd.read_csv(URL1, sep=';')
    df_sec = pd.read_csv(URL2, sep=',')
    df_thr = pd.read_csv(URL3, sep=',')

    df_main=df_main.merge(df_sec,how="left",left_on='ID',right_on='game_id')
    df_main=df_main.merge(df_thr,how="left",left_on='ID',right_on='id')

    return df_main

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Choix des colonnes ,nettoyage avec la fonction _clean_data et ajout de colonnes normalisées.

    La fonction _clean_data supprime les lignes où l'ID BGG est vide, rempli les valeurs vide par des 0
    et change le type des colonnes

    La fonction normalize ajoute des colonnes de valeurs normalisées à partir des colonnes présentes dans le dataframe.

    Args:
        df (pd.DataFrame): Dataframe brut

    Returns:
        pd.DataFrame: Datagreme propre
    """
    df_main=df[COLUMNS]

    df_main=_clean_data(df_main, OBJECT_FLOAT, FILL_NA, NEW_TYPE)

    for col in COLUMNS_TO_NORMALIZE:
        df_main=normalize(df_main, col, col + " n")

    return df_main
    