import pandas as pd


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


def normalize(df: pd.DataFrame, column: str, new_column: str):
    """Normaliser les données d'une colonne

    Args:
        df (pd.DataFrame): df propre
        column (str): colonne à normaliser
        new_column (str): nouvelle colonne

    Returns:
        _type_: df propre
    """
    df[new_column] = (df[column] - df[column].mean()) / (df[column].std())
    return df