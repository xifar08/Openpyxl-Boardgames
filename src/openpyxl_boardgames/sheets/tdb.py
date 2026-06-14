from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl_boardgames.config import (
    SHEET_TDB,
    TITLE_TDB,
    TITLE_CELLS_TDB,
    TITLE_MERGE_TDB
    )


def build_tdb_sheet(wb: Workbook) -> None:
    """_summary_

    Args:
        wb (Workbook): _description_
    """

    ws_tdb=wb.create_sheet(SHEET_TDB)

    ws_tdb.sheet_view.showGridLines = False

    add_titles(ws_tdb)


def add_titles(ws: Worksheet)-> None:
    """_summary_

    Args:
        ws (Worksheet): _description_
    """

    ws["E1"]="Physionomie des jeux de sociétés"
    ws.merge_cells(range_string="E1:M3")
    ws["E1"].font = Font(
        name='Calibri',
        size=28,
        bold=True,
        italic=True
    )
    ws["E1"].alignment = Alignment(horizontal='center', vertical='center',wrapText=True)

    for title, col, merge in zip(TITLE_TDB, TITLE_CELLS_TDB, TITLE_MERGE_TDB):
        ws[col]=title
        ws.merge_cells(range_string=merge)
        ws[col].font = Font(
            name='Calibri',
            size=14,
            bold=True
        )
        ws[col].alignment = Alignment(horizontal='center', vertical='center',wrapText=True)


    # ws_tdb.merge_cells(range_string="B8:C9")
    # ws_tdb.merge_cells(range_string="E8:F9")
    # ws_tdb.merge_cells(range_string="H8:I9")
    # ws_tdb.merge_cells(range_string="K8:L9")
    # ws_tdb.merge_cells(range_string="N8:O9")
    # print("Mise en forme des KPIs")