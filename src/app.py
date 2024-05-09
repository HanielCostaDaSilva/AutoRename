import os
from openpyxl import load_workbook
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
    
    situacao = row["SITUACAO"]
    matricula = row["MATRICULA"]
    nome = row["NOME"] 

    #Caso não esteja PENDENTE a situação do funcionario
    if situacao != fdf.situacao_lista[0]:
        print(f"foi pulado o funcionário: {nome}\n Matricula:{matricula}, Possuindo a seguinte situação: {situacao}")
        continue
    
    continue_autom = autom.confirm(f"Deseja realizar o cadastro de \n {row['NOME']}?")
    
    if not continue_autom:
        if autom.confirm(f"Encerrar o programa? "): break
    

    #MOVE O CURSOR PARA O INPUT DA MATRICULA
    autom.moveCursor(moves[0][0], moves[0][1])
    autom.click(3)
    
    #ESCREVA O INPUT DA MATRICULA
    autom.write(matricula)
    
    #CONFIRME
    autom.pressEnter()
    
    #PERGUNTE SE FOI POSSÍVEL ACESSAR O PERFIL DO FUNCIONÁRIO
    problema_ocorrido=""
    
    while problema_ocorrido in ["","CANCEL"]:
    
        if autom.custom_message("Conseguiu Acesso ao Portal do Funcionario? ",[ "sim", "não"]) == "sim":
            
            problema_ocorrido = autom.custom_message("Informe qual o problema: ", problems + ["ADICIONAR"])
            
            if problema_ocorrido == "ADICIONAR":
                problema_ocorrido = autom.prompt("Adicione um novo erro na lista: ")
                
                if problema_ocorrido not in problems: 
                    problems.append(problema_ocorrido)
                
        else:
            break
  
fdf.save_func_situacao_df()     
# Abre o arquivo Excel
actual_directory = os.path.dirname(os.path.realpath(__file__))

final_file_path = os.path.join(actual_directory,"..","data","final_register")

workbook = load_workbook(final_file_path)
sheet = workbook.active

# Percorre o DataFrame e aplica a formatação às células correspondentes
for index, row in fdf.func_situacao_df.iterrows():
    situacao = row["SITUACAO"]
    matricula = row["MATRICULA"]
    nome = row["NOME"] 
    
    # Verifica a situação e aplica a formatação adequada
    if situacao == "RENOMEADO":
        tf.paintRow(sheet, index + 2, yellow)  # +2 para ajustar o índice da linha
    elif situacao == "NAO NECESSARIO":
        tf.paintRow(sheet, index + 2, orange)
    elif situacao == "ERRO":
        tf.paintRow(sheet, index + 2, green)

# Salva as alterações no arquivo Excel
workbook.save(final_file_path)

autom.alert("Programa encerrado")