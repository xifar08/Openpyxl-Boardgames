# from openpyxl_boardgames.utils import _clean_data, normalize
# from openpyxl_boardgames.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE, HIDDEN
from openpyxl_boardgames.data import load_data, clean_data
from openpyxl_boardgames.config import PATH_TEMPLATE
from openpyxl_boardgames.excel import scatter_chart, bar_chart, data_validation, radar_chart
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl import load_workbook
# import time
import pandas as pd



def main():
    try :

        df_raw=load_data()

        df_main=clean_data(df_raw)

        pd.DataFrame.to_excel(df_main, excel_writer=PATH_TEMPLATE, sheet_name='DATA', index=False)
        print("Données exportées au format .xlsx")

        wb=load_workbook(PATH_TEMPLATE)
        # print("Workbook chargé")
        # wb.create_sheet("TDB", 0)
        # print("Feuille TDB créée")

        # ws_tdb=wb["TDB"]
        # ws_tdb.sheet_view.showGridLines = False

        # ws_tdb["E1"]="Physionomie des jeux de sociétés"
        # ws_tdb.merge_cells(range_string="E1:M3")
        # ws_tdb["E1"].font = Font(
        #     name='Calibri',
        #     size=28,
        #     bold=True,
        #     italic=True
        # )
        # ws_tdb["E1"].alignment = Alignment(horizontal='center', vertical='center',wrapText=True)

        # ws_tdb["B5"]="Filtre famille"
        # ws_tdb.merge_cells(range_string="B5:C7")
        # ws_tdb["B5"].font = Font(
        #     name='Calibri',
        #     size=18,
        #     bold=True
        # )
        # ws_tdb["B5"].alignment = Alignment(horizontal='center', vertical='center',wrapText=True)

        # ws_tdb["E5"]="Nombre de jeux"
        # ws_tdb.merge_cells(range_string="E5:F7")
        # ws_tdb["E5"].font = Font(
        #     name='Calibri',
        #     size=16,
        #     bold=True
        # )
        # ws_tdb["E5"].alignment = Alignment(horizontal='center', vertical='center',wrapText=True)


        # ws_tdb["H5"]="Proportion"
        # ws_tdb.merge_cells(range_string="H5:I7")
        # ws_tdb["H5"].font = Font(
        #     name='Calibri',
        #     size=18,
        #     bold=True
        # )
        # ws_tdb["H5"].alignment = Alignment(horizontal='center', vertical='center',wrapText=True)

        # ws_tdb["K5"]="Rang BGG médian"
        # ws_tdb.merge_cells(range_string="K5:L7")
        # ws_tdb["K5"].font = Font(
        #     name='Calibri',
        #     size=18,
        #     bold=True
        # )
        # ws_tdb["K5"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        # ws_tdb["N5"]="Année moyenne de parution"
        # ws_tdb.merge_cells(range_string="N5:O7")
        # ws_tdb["N5"].font = Font(
        #     name='Calibri',
        #     size=14,
        #     bold=True
        # )
        # ws_tdb["N5"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        # ws_tdb.merge_cells(range_string="B8:C9")
        # ws_tdb.merge_cells(range_string="E8:F9")
        # ws_tdb.merge_cells(range_string="H8:I9")
        # ws_tdb.merge_cells(range_string="K8:L9")
        # ws_tdb.merge_cells(range_string="N8:O9")
        # print("Mise en forme des KPIs")

        # data_validation(wb=wb,worksheet_data=wb["DATA"],worksheet_tdb=wb["TDB"],col_data=14,col_tdb=26,data_title="Liste_domaine",where_tdb="B8",save_as="output/test/test.xlsx")

        # jeux = dict(df_main["Domains"].value_counts())
        # longueur_max=max(jeux.values())
        # formula = '=_xlfn.FILTER(DATA!A:AH,DATA!N:N=TDB!B8,"")'
        # ws_tdb["AA1"] = ArrayFormula(f"AA1:BH{longueur_max}", formula)
        # print("Ajout des données de DATA dans TDB pour être filtrées")

        # scatter_chart(wb,wb["TDB"],wb["TDB"],where="A12",col_x=37,col_y=35,min_row=1,max_row=1000,save_as="output/test/test.xlsx")
        # print("Scatter plot tracé")

        # bar_chart(wb,wb["TDB"],wb["TDB"],where="J12",col_x=28,col_y=38,min_row=1,max_row=10,save_as="output/test/test.xlsx")
        # print("Bar plot tracé")


        # ws_tdb["B8"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        # ws_tdb["E8"]=f'={longueur_max}-COUNTIF(AG:AG,"<>"&NA())'
        # ws_tdb["E8"].font = Font(
        #     name='Calibri',
        #     size=20,
        #     bold=True
        # )
        # ws_tdb["E8"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        # ws_tdb["H8"]="=E8/23327"
        # ws_tdb["H8"].number_format = '0.0%'
        # ws_tdb["H8"].font = Font(
        #     name='Calibri',
        #     size=20,
        #     bold=True
        # )
        # ws_tdb["H8"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        # ws_tdb["K8"]='=_xlfn.AGGREGATE(12,6,AJ:AJ)'
        # ws_tdb["K8"].font = Font(
        #     name='Calibri',
        #     size=20,
        #     bold=True
        # )
        # ws_tdb["K8"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

        # ws_tdb["N8"]='=_xlfn.AGGREGATE(12,6,AC:AC)'
        # ws_tdb["N8"].font = Font(
        #     name='Calibri',
        #     size=20,
        #     bold=True
        # )
        # ws_tdb["N8"].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
        # print("KPI ajoutés")

        


        # wb.create_sheet("TABLEAU",1)
        # ws_tab=wb["TABLEAU"]
        # ws_tab["A1"]="Jeu"
        # ws_tab["A2"]=ArrayFormula("A2:A3030","=_xlfn.UNIQUE(TDB!AB:AB)")
        # ws_tab["B1"]="Année de publication"
        # ws_tab["B2"]=ArrayFormula("B2:B3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AC:AC)")
        # ws_tab["C1"]="Famille"
        # ws_tab["C2"]=ArrayFormula("C2:C3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AN:AN)")
        # ws_tab["D1"]="Rang BGG"
        # ws_tab["D2"]=ArrayFormula("D2:D3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AJ:AJ)")
        # ws_tab["E1"]="Complexité"
        # ws_tab["E2"]=ArrayFormula("E2:E3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AK:AK)")
        # ws_tab["F1"]="Min joueur"
        # ws_tab["F2"]=ArrayFormula("F2:F3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AD:AD)")
        # ws_tab["G1"]="Max joueur"
        # ws_tab["G2"]=ArrayFormula("G2:G3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AE:AE)")
        # ws_tab["H1"]="Temps de jeu"
        # ws_tab["H2"]=ArrayFormula("H2:H3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AF:AF)")
        # ws_tab["I1"]="Game designer"
        # ws_tab["I2"]=ArrayFormula("I2:I3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AP:AP)")
        # ws_tab["J1"]="Artiste"
        # ws_tab["J2"]=ArrayFormula("J2:J3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AO:AO)")
        # ws_tab["K1"]="Editeur"
        # ws_tab["K2"]=ArrayFormula("K2:K3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!AQ:AQ)")
        # ws_tab["L1"]="Note normalisée"
        # ws_tab["L2"]=ArrayFormula("L2:L3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!BA:BA)")
        # ws_tab["M1"]="Complexité normalisée"
        # ws_tab["M2"]=ArrayFormula("M2:M3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!BB:BB)")
        # ws_tab["N1"]="Min joueur normalisé"
        # ws_tab["N2"]=ArrayFormula("N2:N3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!BE:BE)")
        # ws_tab["O1"]="Max joueur normalisé"
        # ws_tab["O2"]=ArrayFormula("O2:O3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!BF:BF)")
        # ws_tab["P1"]="Temps de partie normalisé"
        # ws_tab["P2"]=ArrayFormula("P2:P3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!BG:BG)")
        # ws_tab["Q1"]="Age min normalisé"
        # ws_tab["Q2"]=ArrayFormula("Q2:Q3030","=_xlfn.XLOOKUP(A2:A3030,TDB!AB:AB,TDB!BH:BH)")
        # print("Feuille TABLEAU ajoutée")

        # filters = ws_tab.auto_filter
        # filters.ref = "A:Q"

        # # ws_tdb["W1"]="Indicateurs"
        # ws_tdb["X1"]="Profil de la famille"
        # ws_tdb["Y1"]="Profil du jeu le mieux classé de la famille"
        # ws_tdb["W2"]="Note moyenne normalisée"
        # ws_tdb["X2"]='=_xlfn.AGGREGATE(1,6,BA:BA)'
        # ws_tdb["Y2"]='=TABLEAU!L2'
        # ws_tdb["W3"]="Complexité moyenne normalisée"
        # ws_tdb["X3"]='=_xlfn.AGGREGATE(1,6,BB:BB)'
        # ws_tdb["Y3"]='=TABLEAU!M2'
        # ws_tdb["W4"]="Nombre moyen minimal de joueurs normalisé"
        # ws_tdb["X4"]='=_xlfn.AGGREGATE(1,6,BE:BE)'
        # ws_tdb["Y4"]='=TABLEAU!N2'
        # ws_tdb["W5"]="Nombre moyen maximal de joueurs normalisé"
        # ws_tdb["X5"]='=_xlfn.AGGREGATE(1,6,BF:BF)'
        # ws_tdb["Y5"]='=TABLEAU!O2'
        # ws_tdb["W6"]="Temps de partie moyen normalisé"
        # ws_tdb["X6"]='=_xlfn.AGGREGATE(1,6,BG:BG)'
        # ws_tdb["Y6"]='=TABLEAU!P2'
        # ws_tdb["W7"]="Age minimal moyen normalisé"
        # ws_tdb["X7"]='=_xlfn.AGGREGATE(1,6,BH:BH)'
        # ws_tdb["Y7"]='=TABLEAU!Q2'
        # wb.save("output/test/test.xlsx")

        # radar_chart(wb,wb["TDB"],wb["TDB"],where="A37",min_row=1,max_row=7,min_col=23,save_as="output/test/test.xlsx")
        # print("Radar chart tracé")




        #les graphiques ne fonctionnent plus si je cache les colonnes
        # for col in HIDDEN:
        #     ws_tdb.column_dimensions[col].hidden = True
        # print("Colonnes masquées")

        # wb.save("output/test/test.xlsx")

    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
