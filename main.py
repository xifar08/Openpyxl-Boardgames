from src.utils import get_data, join_df, choose_columns, clean_data, save_data
from src.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE
from src.excel import load_wb, create_ws, scatter_chart, copy_columns, bar_chart, data_validation
from openpyxl.worksheet.formula import ArrayFormula
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
        print("Jointure effectuée")

        df_main=choose_columns(df_main, COLUMNS)

        df_main=clean_data(df_main, OBJECT_FLOAT, FILL_NA, NEW_TYPE)
        print("Données nettoyées")

        save_data(df_main)
        print("Données exportées au format .xlsx")

        wb=load_wb("template.xlsx")
        print("Workbook chargé")

        create_ws(wb,"TDB","test.xlsx")
        print("Feuille TDB créée")

        create_ws(wb,"SCATTER","test.xlsx")
        print("Feuille SCATTER créée")

        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=10, col_destination=1,min_row=1,save_as="test.xlsx")
        print("Données BGG rank copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=9, col_destination=2,min_row=1,save_as="test.xlsx")
        print("Données Rating copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=11, col_destination=3,min_row=1,save_as="test.xlsx")
        print("Données Complexity copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=8, col_destination=4,min_row=1,save_as="test.xlsx")
        print("Données User rated copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=14, col_destination=5,min_row=1,save_as="test.xlsx")
        print("Données Domains copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=2, col_destination=6,min_row=1,save_as="test.xlsx")
        print("Données Name copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=12, col_destination=7,min_row=1,save_as="test.xlsx")
        print("Données Owned User copiées dans la feuille Scatter")

        data_validation(wb=wb,worksheet_data=wb["SCATTER"],worksheet_tdb=wb["TDB"],col_data=5,col_tdb=26,data_title="Liste_domaine",where_tdb="A2",save_as="test.xlsx")
        ws_tdb=wb["TDB"]
        ws_tdb["A1"] = "Filtre Famille"
        print("Liste des familles ajoutée.")

        formula = '=_xlfn.FILTER(SCATTER!A:G,SCATTER!E:E=TDB!A2,"")'
        ws_tdb["AA1"] = ArrayFormula("AA:AG", formula)
        print("Ajout des données de SCATTER dans TDB pour être filtrées")

        scatter_chart(wb,wb["TDB"],wb["TDB"],where="A10",col_x=29,col_y=28,min_row=1,max_row=1000,save_as="test.xlsx")
        print("Scatter plot tracé")

        bar_chart(wb,wb["TDB"],wb["TDB"],where="M10",col_x=32,col_y=33,min_row=1,max_row=10,save_as="test.xlsx")
        print("Bar plot tracé")

    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
