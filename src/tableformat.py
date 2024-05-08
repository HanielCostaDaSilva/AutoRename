from unidecode import unidecode
from openpyxl.styles import PatternFill

def format_cell(cell_value:str)-> str:
    cell_value = unidecode(cell_value).upper()
    return cell_value

def paintRow(row, collor:PatternFill):
    #pinta todas as colunas de uma determinada linha 
    for cell in row:
        cell.fill = collor