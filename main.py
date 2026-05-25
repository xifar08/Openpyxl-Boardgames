from src.utils import get_data, join_df, choose_columns, clean_data, save_data
from src.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE
from src.excel import load_wb, create_ws, scatter_chart, copy_columns
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

        wb=load_wb("template.xlsx")
        print("Workbook chargé")

        create_ws(wb,"TDB","test.xlsx")
        print("Feuille TDB créée")

        create_ws(wb,"SCATTER","test.xlsx")
        print("Feuille SCATTER créée")

        copy_columns(wb,wb["DATA"],wb["SCATTER"], col=9,min_row=2, save_as="test.xlsx")

        #il faut que je copie les données qui m'interessent dans une autre feuille


        #scatter_chart(wb["TDB"],wb["DATA"],)
    
    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
