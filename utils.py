import pandas as pd


def get_data(url: str, sep: str) -> pd.DataFrame:
    return pd.read_csv(url, sep=sep)
