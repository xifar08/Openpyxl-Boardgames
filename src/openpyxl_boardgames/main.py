from openpyxl import load_workbook
from openpyxl_boardgames.data import load_data, clean_data
from openpyxl_boardgames.dashboard import build_workbook
from openpyxl_boardgames.config import PATH_TEMPLATE, SHEET_DATA
import pandas as pd
# from openpyxl_boardgames.utils import _clean_data, normalize
# from openpyxl_boardgames.config import URL1, URL2, URL3, COLUMNS, OBJECT_FLOAT, FILL_NA, NEW_TYPE, HIDDEN
# from openpyxl_boardgames.excel import scatter_chart, bar_chart, data_validation, radar_chart
# from openpyxl.worksheet.formula import ArrayFormula
# import time




def main():
    try :

        df_raw=load_data()

        df_main=clean_data(df_raw)

        pd.DataFrame.to_excel(df_main, excel_writer=PATH_TEMPLATE, sheet_name=SHEET_DATA, index=False)
        print("Données exportées au format .xlsx")

        wb_in=load_workbook(PATH_TEMPLATE)

        wb_out=build_workbook(wb_in)

        wb_out.save("output/test/refacto.xlsx")

        

    except Exception as e:
        print(f"Erreur : {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
