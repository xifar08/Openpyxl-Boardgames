from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl_boardgames.config import (
    SHEET_TABLEAU,
    TITLE_TABLEAU,
    TITLE_CELLS_TABLEAU,
    KPI_CELLS_TABLEAU,
    KPI_FORMULA_TABLEAU,
    KPI_FORMULA_RANGE
    )

def build_tdb_tableau(wb: Workbook) -> None:
    """_summary_

    Args:
        wb (Workbook): _description_
    """
    ws_tab=wb.create_sheet(SHEET_TABLEAU,1)

    add_titles(ws=ws_tab)

    add_kpis(ws=ws_tab)


def add_titles(ws: Worksheet)-> None :
    """_summary_

    Args:
        ws (Worksheet): _description_
    """
    for title, col in zip(TITLE_TABLEAU, TITLE_CELLS_TABLEAU):
        ws[col]=title
        ws[col].font = Font(
            name='Calibri',
            size=11,
            bold=True
        )


def add_kpis(ws: Worksheet)-> None :
    """_summary_

    Args:
        ws (Worksheet): _description_
    """
    for formula, col, rg in zip(KPI_FORMULA_TABLEAU, KPI_CELLS_TABLEAU, KPI_FORMULA_RANGE):
        ws[col]=ArrayFormula(rg, formula)