from openpyxl import Workbook

from openpyxl_boardgames.sheets.tdb import build_tdb_sheet
from openpyxl_boardgames.sheets.tableau import build_tdb_tableau

def build_workbook(wb: Workbook)->Workbook:
    """_summary_

    Args:
        wb (Workbook): _description_

    Returns:
        Workbook: _description_
    """
    build_tdb_sheet(wb=wb)
    build_tdb_tableau(wb=wb)

    return wb