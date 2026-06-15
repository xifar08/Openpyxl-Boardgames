"""Point d'entrée CLI : génération du reporting Excel"""

from openpyxl_boardgames.data import load_data, clean_data
from openpyxl_boardgames.dashboard import build_workbook
from openpyxl_boardgames.config import PATH_TEMPLATE, SHEET_DATA, OUTPUT_PATH

import pandas as pd
from openpyxl import load_workbook




def main():
    """Génère le classeur Excel de reproting à partir des différentes tables de données.

    Orchestre le pipeline complet : chargement et nettoyage des données, enregistrement 
    en classeur intermédiaire et construction des onglets. Sauvegarde dans OUTPUTH_PATH
    """

    df_raw=load_data()
    print("Données chargées")

    df_main=clean_data(df_raw)
    print("Données nettoyées")

    PATH_TEMPLATE.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame.to_excel(df_main, excel_writer=PATH_TEMPLATE, sheet_name=SHEET_DATA, index=False)
    print("Données exportées au format .xlsx")

    wb_in=load_workbook(PATH_TEMPLATE)
    print("Reporting initialisé")

    wb_out=build_workbook(wb_in)
    print("Reporting construit")

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(OUTPUT_PATH)
    print("Reporting exporté avec succés")


if __name__ == "__main__":
    main()
