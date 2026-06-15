"""Orchestration : assemblage du classeur Excel complet en mémoire"""

from openpyxl import Workbook

from openpyxl_boardgames.sheets.tdb import build_tdb_sheet
from openpyxl_boardgames.sheets.tableau import build_tdb_tableau

def build_workbook(wb: Workbook)->Workbook:
    """Construit le classeur Excel complet en mémoire. 

    L'onglet TDB est crée en premier et l'onglet TABLEAU ensuite.

    Args:
        wb (Workbook): Workbook avec uniquement les données brutes

    Returns:
        Workbook: Workbook final
    """
    build_tdb_sheet(wb=wb)
    build_tdb_tableau(wb=wb)

    return wb