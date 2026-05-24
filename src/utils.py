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
        pd.DataFrame: jointure des deux dataframes
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


def clean_data(df: pd.DataFrame, object_float: list, fillna: dict, type_dict: dict ) -> pd.DataFrame:
    """Nettoyage du df en supprimant les IDs nuls et en changeant les types des colonnes

    Args:
        df (pd.DataFrame): df à nettoyer
        object_float (list): liste des colonnes où il faut remplacer le '.' par ','
        fillna (dict): dictionnaire avec pour clés les colonnes et en valeur la valeur à utiliser pour remplir les cellules NA
        type_dict (dict): dictionnaire avec pour clés les colonnes et en valeur le type à appliquer

    Returns:
        pd.DataFrame: df clean
    """
    df=df.dropna(subset='ID')
    for k in object_float:
        df[k]=df[k].str.replace(",", ".")
    df=df.fillna(fillna)
    df=df.astype(type_dict)
    return df


def save_data(df: pd.DataFrame):
    return pd.DataFrame.to_excel(df, excel_writer='template.xlsx', sheet_name='DATA', index=False)