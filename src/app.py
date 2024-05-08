from openpyxl.styles import PatternFill
import tableformat as tf
from autom import autom
from dataframes import funcionario as fdf



def get_problem()->str:
    pass

#== == == Variables
#== == moves_to_form
moves=[ (2181,384), (2694,486)]

#== == possibles_problem
problems=["MATRICULA INEXISTENTE"]

# == == Collors
#AMARELO quando o funcionário foi devidamente renoemado 
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

#LARANJA quando não foi necessário renomeiar
orange = PatternFill(start_color="EC9006", end_color="EC9006", fill_type="solid")

#VERDE quando ocorreu um erro ao tentar inserir o funcionário
green = PatternFill(start_color="40DE47", end_color="40DE47", fill_type="solid")

continue_autom= False

for index,row in fdf.func_situacao_df.iterrows():
    continue_autom = autom.confirm(f"Deseja realizar o cadastro de \n {row["NOME"]}?")
    
    if not continue_autom:
        if autom.confirm(f"Encerrar o programa? "): break
    
    matricula = row["MATRICULA"]
    nome = row["NOME"] 
    #MOVE O CURSOR PARA A ENTRADA DA MATRICULA
    autom.moveCursor(moves[0][0], moves[0][1])
    autom.click(3)
    
    #ESCREVA A ENTRADA DA MATRICULA
    autom.write(matricula)
    
    #CONFIRME
    autom.pressEnter()
    
    #PERGUNTE SE FOI POSSÍVEL ACESSAR O PERFIL DO FUNCIONÁRIO
    problema_ocorrido=""
    
    while problema_ocorrido in ["","CANCEL"]:
    
        if autom.custom_message("Aconteceu algum erro? ",[ "sim", "não"]) == "sim":
            
            problema_ocorrido = autom.custom_message("Informe qual o problema: ", problems + ["ADICIONAR"])
            
            if problema_ocorrido == "ADICIONAR":
                problema_ocorrido = autom.prompt("Adicione um novo erro na lista: ")
                
                if problema_ocorrido not in problems: 
                    problems.append(problema_ocorrido)
            
        else:
            break
              
autom.alert("Programa encerrado")