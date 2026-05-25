from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.chart import (
    LineChart,
    Reference,
)

def load_wb(path: str) -> Workbook:
    return load_workbook(filename = path)


def create_ws(wb: Workbook, sheet_name: str, save_as: str):
    ws = wb.create_sheet(sheet_name, 0)
    wb.save(save_as)
    # je ne sais pas s'il vaut pas mieux retrouner le wb pour pouvoir l'utiliser au global
    #return ws

def copy_columns(wb: Workbook, ws_data: Worksheet, ws_destination: Worksheet, col: int, min_row: int, save_as: str):
    for row in ws_data.iter_rows(min_row=min_row, max_row=ws_data.max_row):
        val=row[col-1].value
        ws_destination.append([val])
    wb.save(save_as)



def scatter_chart(wb: Workbook, worksheet_chart: Worksheet, worksheet_data: Worksheet, where: str, min_col: int, max_col: int, min_row: int, max_row: int, save_as: str):
    c1 = LineChart()
    c1.title = "Complexité vs Note utilisateur"
    c1.style = 13
    c1.y_axis.title = 'Note utilisateur'
    c1.x_axis.title = 'Complexité'

    data = Reference(worksheet_data, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
    c1.add_data(data, titles_from_data=True)

    s1 = c1.series[0]
    s1.graphicalProperties.line.solidFill = "00AAAA"
    s1.graphicalProperties.line.dashStyle = "sysDot"
    s1.graphicalProperties.line.width = 100050 # width in EMUs

    worksheet_chart.add_chart(c1, where)

    wb.save(save_as)

    # a voir si  vaut mieux pas retourner le workbook ou rien retourner
    #return worksheet_chart
