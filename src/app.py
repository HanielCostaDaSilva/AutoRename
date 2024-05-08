from openpyxl.styles import PatternFill
import tableformat as tf
from autom import autom
from dataframes import funcionario

#== == == Variables
# == == Collors

#AMARELO quando o funcionário foi devidamente renoemado 
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

#LARANJA quando não foi necessário renomeiar
orange = PatternFill(start_color="EC9006", end_color="EC9006", fill_type="solid")

#VERDE quando ocorreu um erro ao tentar inserir o funcionário
green = PatternFill(start_color="40DE47", end_color="40DE47", fill_type="solid")
