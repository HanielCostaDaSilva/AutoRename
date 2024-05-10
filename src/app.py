import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tableformat as tf
from autom import autom
from dataframes import funcionario as fdf

def get_problem(text:str="")->str:
    """
    Cria uma janeja perguntando ao usuario se ocorreu algum problema ao tentar execultar alguma operação.
    Se ele escolher "sim", o programa então criará uma janela perguntando a ele qual o problema que aconteceu.

    Returns:
        str: o problema que o usuário relatou, uma string vazia caso nenhum problema aconteceu
    """
    ocorreu_problema = True
    problema= ""
    
    while ocorreu_problema ==True and problema =="":    
        #pergunta se aconteceu algum problema
        ocorreu_problema = autom.custom_message(f"{text}\nOcorreu algum problema? ",["sim", "não"]) == "sim"
        #Caso não tenha ocorrido algum problema

        if not ocorreu_problema:
            break
        
        #caso tenha acontecido algum problema, pedimos que ele informe qual foi o problema que aconteceu
        escolha = autom.custom_message("Informe qual problema aconteceu: ", problems + ["ADICIONAR","LIMPAR"])
    
        if escolha == "ADICIONAR":
            problema = autom.prompt("Adicione um novo erro na lista: ")
            
            if escolha != "" and escolha not in problems:
                problems.append(problema)
        
        elif escolha == "LIMPAR": #caso ele tenha escolhido limpar os possíveis problemas
            problems.clear()
            problems.append("MATRICULA INEXISTENTE")
            continue        
        
        elif escolha =="": #Caso ele tenha fechado a janela, refaça a pergunta
            continue
        
        else:
            problema = escolha     
    
    return problema

#== == == Variables
#== == moves_to_form
moves=[ (2181,384), (2694,486)]

#== == possibles_problem
problems=["MATRICULA INEXISTENTE"]

# == == Collors
#AMARELO quando o funcionário foi devidamente renoemado 
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

#LARANJA quando não foi necessário renomear
orange = PatternFill(start_color="EC9006", end_color="EC9006", fill_type="solid")

#VERDE quando ocorreu um erro ao tentar renomear o funcionário
green = PatternFill(start_color="40DE47", end_color="40DE47", fill_type="solid")

continue_autom= False

autom.alert("Programa Inciado. Por favor, conecte-se a sua conta")

for index,row in fdf.func_situacao_df.iterrows():
    
    situacao = row["SITUACAO"]
    matricula = row["MATRICULA"]
    nome = row["NOME"] 

    #Caso não esteja PENDENTE a situação do funcionario
    if situacao != fdf.situacao_lista[0]:
        print(f"foi pulado o funcionário: {nome}\n Matricula:{matricula}, Situação: {situacao}")
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
    problem = get_problem(f"Você está tentando acessar: {nome}")
    
    if problem != "":
        situacao= problem
    
    #caso tenha conseguido acessar ao portal do funcionário
    else:
        #Continue com a automação
        autom.moveCursor(moves[1][0], moves[1][1]) #mova o cursor para o input de nome
        autom.click(2)
        autom.write(nome)
        autom.pressEnter(2)
       
        problem = get_problem(f"Você está tentando alterar: {nome}")
        
        if problem != "":
            #foi detectado algum erro ao tentar renomear o usuário
            situacao= problem
        else:
            situacao = fdf.situacao_lista[-1]     
            autom.pressEnter(1)
            
    row["SITUACAO"]=situacao
    print(row)

fdf.save_func_situacao_df()     

# Carregar o arquivo Excel mesclado
workbook = load_workbook(fdf.func_situacao_path)
ws = workbook.active

for row in ws.iter_rows(min_row=2, max_row = len(fdf.func_situacao_df) + 1, min_col=1, max_col=len(fdf.func_situacao_df.columns)):

    situacao_func= row[fdf.func_situacao_df.columns.get_loc("SITUACAO")].value

    if situacao_func ==fdf.situacao_lista[-1]: #ALTERADO
        tf.paintRow(row,yellow)
    
    elif situacao_func == fdf.situacao_lista[0]:
        continue
    
    else: #Ocorreu algum problema
        tf.paintRow(row,green)
    
# Salva as alterações no arquivo Excel
workbook.save(fdf.func_situacao_path)

autom.alert("PROGRAMA FINALIZADO")