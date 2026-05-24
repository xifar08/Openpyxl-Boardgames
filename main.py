from src.utils import get_data, join_df, choose_columns, clean_data, save_data
from src.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE
from src.excel import create_ws


def main():
    df_main = get_data(URL1, sep=';')
    df_sec = get_data(URL2, sep=',')
    df_thr = get_data(URL3, sep=',')
    print("Données récupérées")

    df_main=join_df(df_left=df_main, df_right=df_sec,key_left='ID',key_right='game_id',how_join='left')
    df_main=join_df(df_left=df_main, df_right=df_thr,key_left='ID',key_right='id',how_join='left')
    print("Jointure effectuée")

    df_main=choose_columns(df_main, COLUMNS)

    df_main=clean_data(df_main, OBJECT_FLOAT, FILL_NA, NEW_TYPE)
    print("Données nettoyées")

    save_data(df_main)
    print("Données exportées au format .xlsx")

    create_ws("../template.xlsx","TDB","test.xlsx")
    print("Feuille TDB créée")


if __name__ == "__main__":
    main()
