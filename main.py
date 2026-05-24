from src.utils import get_data, join_df, choose_columns, clean_data, save_data
from src.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE
from src.excel import create_ws
import time


def main():
    try :
        t = time.perf_counter()

        df_main = get_data(URL1, sep=';')
        df_sec = get_data(URL2, sep=',')
        df_thr = get_data(URL3, sep=',')
        print(f"Données récupérées en {time.perf_counter()-t:.2f}s")

        df_main=join_df(df_left=df_main, df_right=df_sec,key_left='ID',key_right='game_id',how_join='left')
        df_main=join_df(df_left=df_main, df_right=df_thr,key_left='ID',key_right='id',how_join='left')
        print(f"Jointure effectuée en {time.perf_counter()-t:.2f}s")

        df_main=choose_columns(df_main, COLUMNS)

        df_main=clean_data(df_main, OBJECT_FLOAT, FILL_NA, NEW_TYPE)
        print(f"Données nettoyées en {time.perf_counter()-t:.2f}s")

        save_data(df_main)
        print(f"Données exportées au format .xlsx en {time.perf_counter()-t:.2f}s")

        ws=create_ws("template.xlsx","TDB","test.xlsx")
        print(f"Feuille TDB créée en {time.perf_counter()-t:.2f}s")
    
    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
