from openpyxl_boardgames.utils import get_data, join_df, choose_columns, clean_data, save_data, normalize
from openpyxl_boardgames.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE
from openpyxl_boardgames.excel import load_wb, create_ws, scatter_chart, copy_columns, bar_chart, data_validation
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

        df_main=normalize(df_main,"Rating Average","Rating Average n")
        df_main=normalize(df_main,"Complexity Average","Complexity Average n")
        df_main=normalize(df_main,"Users Rated", "Users Rated n")
        df_main=normalize(df_main,"Owned Users", "Owned Users n")
        df_main=normalize(df_main,"Min Players", "Min Players n")
        df_main=normalize(df_main,"Max Players", "Max Players n")
        df_main=normalize(df_main,"Play Time", "Play Time n")
        df_main=normalize(df_main,"Min Age", "Min Age n")
        print("Données normalisées")

        save_data(df_main)
        print("Données exportées au format .xlsx")

        wb=load_wb("template.xlsx")
        print("Workbook chargé")

        create_ws(wb,"TDB","output/test/test.xlsx")
        print("Feuille TDB créée")


        data_validation(wb=wb,worksheet_data=wb["DATA"],worksheet_tdb=wb["TDB"],col_data=14,col_tdb=26,data_title="Liste_domaine",where_tdb="A2",save_as="output/test/test.xlsx")
        ws_tdb=wb["TDB"]
        ws_tdb["A1"] = "Filtre Famille"
        print("Liste des familles ajoutée.")

        jeux = dict(df_main["Domains"].value_counts())
        longueur_max=max(jeux.values())
        formula = '=_xlfn.FILTER(DATA!A:AH,DATA!N:N=TDB!A2,"")'
        ws_tdb["AA1"] = ArrayFormula(f"AA1:BH{longueur_max}", formula)
        print("Ajout des données de DATA dans TDB pour être filtrées")

        scatter_chart(wb,wb["TDB"],wb["TDB"],where="A10",col_x=37,col_y=35,min_row=1,max_row=1000,save_as="output/test/test.xlsx")
        print("Scatter plot tracé")

        bar_chart(wb,wb["TDB"],wb["TDB"],where="M10",col_x=28,col_y=38,min_row=1,max_row=10,save_as="output/test/test.xlsx")
        print("Bar plot tracé")

        ws_tdb["B1"]="Nombre de jeux"
        ws_tdb["B2"]=f'={longueur_max}-COUNTIF(AG:AG,"<>"&NA())'
        print("KPI nombre de jeux ajouté")
        ws_tdb["C1"]="Rang BGG médian"
        ws_tdb["C2"]='=_xlfn.AGGREGATE(12,6,AJ:AJ)'
        print("KPI médiane ajouté")

        ws_tdb["A4"]="Note moyenne normalisée"
        ws_tdb["A5"]='=_xlfn.AGGREGATE(1,6,BA:BA)'
        ws_tdb["B4"]="Complexité moyenne normalisée"
        ws_tdb["B5"]='=_xlfn.AGGREGATE(1,6,BB:BB)'
        ws_tdb["C4"]="Nombre moyen minimal de joueurs normalisé"
        ws_tdb["C5"]='=_xlfn.AGGREGATE(1,6,BE:BE)'
        ws_tdb["D4"]="Nombre moyen maximal de joueurs normalisé"
        ws_tdb["D5"]='=_xlfn.AGGREGATE(1,6,BF:BF)'
        ws_tdb["E4"]="Temps de partie moyen normalisé"
        ws_tdb["E5"]='=_xlfn.AGGREGATE(1,6,BG:BG)'
        ws_tdb["F4"]="Age minimal moyen normalisé"
        ws_tdb["F5"]='=_xlfn.AGGREGATE(1,6,BH:BH)'
        wb.save("output/test/feat-5.xlsx")

    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
