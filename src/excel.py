from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

def create_ws(path: str, sheet_name: str, save_as: str) -> Worksheet:
    """Création d'une feuille excel

    Args:
        path (str): fichier excel à charger
        sheet_name (str): nom de la feuille
        save_as (str): fichier excel de sortie

    Returns:
        Worksheet: feuille créée
    """
    wb = load_workbook(filename = path)
    ws = wb.create_sheet(sheet_name, 0)
    wb.save(save_as)
    return ws