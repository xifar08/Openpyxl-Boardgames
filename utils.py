import pandas as pd


def get_data(url: str, sep: str) -> pd.DataFrame:
    """Retourner un dataframe à partir d'un fichier .csv disponible sur le net

    Args:
        url (str): URL du fichier
        sep (str): Séparateur du .csv

    Returns:
        pd.DataFrame: dataframe du .csv
    """
    return pd.read_csv(url, sep=sep)


def join_df(df_left: pd.DataFrame, df_right: pd.DataFrame, key_left: str, key_right: str, how_join: str) -> pd.DataFrame:
    """Joindre 2 dataframes

    Args:
        df_left (pd.DataFrame): df de gauche
        df_right (pd.DataFrame): de de droite
        key_left (str): clé de gauche pour la jointure
        key_right (str): clé de droite pour la jointure
        how_join (str): type de jointure (left, right, inner, etc)

    Returns:
        pd.DataFrame: dataframe
    """
    return df_left.merge(df_right,how=how_join,left_on=key_left,right_on=key_right ) # type: ignore


def choose_columns(df: pd.DataFrame, columns: list) -> pd.DataFrame:
    """Sélections les colonnes d'un df parent pour générer un df enfant

    Args:
        df (pd.DataFrame): df parent
        columns (list): liste des colonnes à sélectionner

    Returns:
        pd.DataFrame: df enfant
    """
    return df[columns]
