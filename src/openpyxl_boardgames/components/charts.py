from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.chart import (
    ScatterChart,
    BubbleChart,
    Reference,
    BarChart,
    RadarChart,
    #existe bien
    Series
)

def scatter_chart( 
        worksheet_chart: Worksheet, 
        worksheet_data: Worksheet, 
        where: str, 
        col_x: int, 
        col_y: int, 
        min_row: int, 
        max_row: int
        ) -> None :
    
    
    c1 = ScatterChart()
    c1.title = "Complexité vs Note utilisateur"
    #c1.style = 5
    c1.legend = None
    c1.y_axis.title = 'Note utilisateur'
    c1.x_axis.title = 'Complexité'
    # c1.width = 20
    c1.height = 12

    x_values = Reference(worksheet_data, min_col=col_x, min_row=min_row, max_row=max_row)
    y_values = Reference(worksheet_data, min_col=col_y, min_row=min_row, max_row=max_row)
    series = Series(y_values, x_values, title_from_data=False)

    series.marker.symbol = "circle" 
    series.marker.size = 3
    series.graphicalProperties.line.noFill = True
    
    c1.series.append(series)

    worksheet_chart.add_chart(c1, where)


def bubble_chart(
        worksheet_chart: Worksheet, 
        worksheet_data: Worksheet, 
        where: str, 
        col_x: int, 
        col_y: int,
        col_size: int, 
        min_row: int, 
        max_row: int
        ) -> None :
    """_summary_

    Args:
        wb (Workbook): _description_
        worksheet_chart (Worksheet): _description_
        worksheet_data (Worksheet): _description_
        where (str): _description_
        col_x (int): _description_
        col_y (int): _description_
        col_size (int): _description_
        min_row (int): _description_
        max_row (int): _description_
        save_as (str): _description_
    """
    c1 = BubbleChart()
    c1.title = "Complexité vs Note utilisateur"
    c1.style = 1
    c1.legend = None
    c1.y_axis.title = 'Note utilisateur'
    c1.x_axis.title = 'Complexité'
    # c1.width = 20
    # c1.height = 15

    x_values = Reference(worksheet_data, min_col=col_x, min_row=min_row, max_row=max_row)
    y_values = Reference(worksheet_data, min_col=col_y, min_row=min_row, max_row=max_row)
    size = Reference(worksheet_data,min_col=col_size, min_row=min_row, max_row=max_row)
    series = Series(y_values, x_values, size, title_from_data=True)
    c1.series.append(series)

    worksheet_chart.add_chart(c1, where)


# def add_filter(wb: Workbook, worksheet: Worksheet, range: str, save_as: str) -> None:
#     """_summary_

#     Args:
#         wb (Workbook): _description_
#         worksheet (Worksheet): _description_
#         range (str): _description_
#         save_as (str): _description_
#     """
#     filters = worksheet.auto_filter
#     # peu importe la ligne de fin mais il faut les bonnes colonnes
#     filters.ref = range
#     wb.save(save_as)



def bar_chart( 
        worksheet_chart: Worksheet, 
        worksheet_data: Worksheet, 
        where: str, 
        col_x: int, 
        col_y: int, 
        min_row: int, 
        max_row: int
        ) -> None :
    """_summary_

    Args:
        worksheet_chart (Worksheet): _description_
        worksheet_data (Worksheet): _description_
        where (str): _description_
        col_x (int): _description_
        col_y (int): _description_
        min_row (int): _description_
        max_row (int): _description_
    """
    
    c1 = BarChart()
    c1.title="Top 10 BGG : Nombre d'unités de jeux possédées"
    c1.type="col"
    c1.legend = None
    c1.y_axis.title = "Nombre d'unités jeux possédés"
    c1.x_axis.title = 'Jeux'
    # c1.width = 20
    c1.height = 12

    x_values = Reference(worksheet_data, min_col=col_x, min_row=min_row, max_row=max_row)
    y_values = Reference(worksheet_data, min_col=col_y, min_row=min_row, max_row=max_row)

    c1.add_data(y_values, titles_from_data=False)
    c1.set_categories(x_values)

    worksheet_chart.add_chart(c1, where)


def radar_chart( 
        worksheet_chart: Worksheet, 
        worksheet_data: Worksheet, 
        where: str, 
        min_row: int,
        max_row: int, 
        min_col: int 
        ) -> None :
    """_summary_

    Args:
        worksheet_chart (Worksheet): _description_
        worksheet_data (Worksheet): _description_
        where (str): _description_
        min_row (int): _description_
        max_row (int): _description_
        min_col (int): _description_
    """
    c1=RadarChart()
    c1.title = "Carte d'identité de la famille de jeux"
    labels=Reference(worksheet_data, min_col=min_col,min_row=min_row+1,max_row=max_row)
    y_values=Reference(worksheet_data, min_col=min_col+1,max_col=min_col+2,min_row=min_row,max_row=max_row)
    c1.add_data(y_values, titles_from_data=True)
    c1.set_categories(labels)
    c1.y_axis.delete=True
    # c1.legend = None
    # c1.width = 20
    c1.height = 12

    worksheet_chart.add_chart(c1, where)