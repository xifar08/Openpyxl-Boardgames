from utils import get_data, join_df, choose_columns
from config import URL1, URL2, URL3, COLUMNS


def main():
    df_main = get_data(URL1, sep=';')
    df_sec = get_data(URL2, sep=',')
    df_thr = get_data(URL3, sep=',')

    df_main=join_df(df_left=df_main, df_right=df_sec,key_left='ID',key_right='game_id',how_join='left')
    df_main=join_df(df_left=df_main, df_right=df_thr,key_left='ID',key_right='id',how_join='left')

    df_main=choose_columns(df_main, COLUMNS)

    print(df_main.columns)
    print(df_main.dtypes)


if __name__ == "__main__":
    main()
