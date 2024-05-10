import pandas as pd
import os

#== == == Variables
# == == Caminhos
# Obter o diretório atual do arquivo de script
actual_directory = os.path.dirname(os.path.realpath(__file__))

func_path = os.path.join(actual_directory,"..","..","data","func.xlsx")

func_situacao_path = os.path.join(actual_directory,"..","..","data","func_log.xlsx")

# == == Lista
situacao_lista = ["PENDENTE","ALTERADO"]
#PENDENTE = ainda não foi checado a situação do Funcionario
#ALTERADO = O nome do funcionário foi alterado  

def save_func_situacao_df():
    func_situacao_df.to_excel(func_situacao_path,index=False)
    

def get_func_original_df(path:str):
    func_df = pd.read_excel(path, dtype=str)
    return func_df

def create_func_situacao_df(func_df:'pd.DataFrame'):
    '''Create a new dataframe with the situations of the employees'''
    func_situacao_df = func_df.copy()
    #Adicionamos a coluna SITUACAO no final da tabela 
    func_situacao_df["SITUACAO"] = situacao_lista[0]
    #Criamos a planilha na pasta data
    func_situacao_df.to_excel(func_situacao_path,index=False)
    
    return func_situacao_df
    
def get_func_situacao(path):  
    # Carregue os cabeçalhos da planilha existente
    func_log_df = pd.read_excel(path, dtype=str)    
    #Preenchemos as colunas vazias de situação por "situacao_lista[0]"
    func_log_df.loc[func_log_df['SITUACAO'] == '', 'SITUACAO'] = situacao_lista[0]    
    return func_log_df

# == == DFs
func_orginal_df =get_func_original_df(func_path)

func_situacao_df = None

if os.path.exists(func_situacao_path):
    # Carregue os cabeçalhos da planilha existente
    func_situacao_df = get_func_situacao(func_situacao_path)
       
else: 
    func_situacao_df = create_func_situacao_df(func_orginal_df)

if __name__ == '__main__':
    print(func_orginal_df)
    print("=="*30)
    print(func_situacao_df)