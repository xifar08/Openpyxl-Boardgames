from src.utils import get_data, join_df, choose_columns, clean_data, save_data
from src.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE
from src.excel import load_wb, create_ws, scatter_chart, copy_columns, bubble_chart, data_validation
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

        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=9, col_destination=1,min_row=1,save_as="test.xlsx")
        print("Données 1ere colonne copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=11, col_destination=2,min_row=1,save_as="test.xlsx")
        print("Données 2eme colonne copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=8, col_destination=3,min_row=1,save_as="test.xlsx")
        print("Données 3eme colonne copiées dans la feuille Scatter")
        copy_columns(wb,wb["DATA"],wb["SCATTER"],col_data=14, col_destination=4,min_row=1,save_as="test.xlsx")
        print("Données 4eme colonne copiées dans la feuille Scatter")

        # bubble_chart(wb,wb["TDB"],wb["SCATTER"],where="A10",col_x=1,col_y=2,col_size=3,min_row=1,max_row=200,save_as="test.xlsx")
        # print("Bubble plot tracé")
        # scatter_chart(wb,wb["TDB"],wb["SCATTER"],where="M10",col_x=1,col_y=2,min_row=1,max_row=200,save_as="test.xlsx")
        # print("Scatter plot tracé")

        # add_filter(wb,wb["SCATTER"],range="A1:D2",save_as="test.xlsx")
        # print("Filtre ajouté")

        data_validation(wb=wb,worksheet_data=wb["SCATTER"],worksheet_tdb=wb["TDB"],col_data=4,col_tdb=26,data_title="Liste_domaine",where_tdb="A2",save_as="test.xlsx")
        print("Liste des domaines ajoutée.")

        formula = '=_xlfn.FILTER(SCATTER!A:D,SCATTER!D:D=TDB!A2,"")'
        ws=wb["TDB"]
        ws["AA1"] = ArrayFormula("AA:AD", formula)
        print("Ajout des données de SCATTER dans TDB pour être filtrées")

        scatter_chart(wb,wb["TDB"],wb["TDB"],where="A10",col_x=28,col_y=27,min_row=1,max_row=1000,save_as="test.xlsx")
        print("Scatter plot tracé")

    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
