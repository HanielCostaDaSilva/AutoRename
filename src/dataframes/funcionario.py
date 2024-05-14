import pandas as pd
import os

#== == == Variables
# == == Caminhos
# Obter o diretório atual do arquivo de script
actual_directory = os.path.dirname(os.path.realpath(__file__))

func_path = os.path.join(actual_directory,"..","..","data","func.xlsx")

func_situacao_path = os.path.join(actual_directory,"..","..","data","func_log.xlsx")

# == == Lista
situacao_lista = ["PENDENTE","INALTERADO","ALTERADO"]
#PENDENTE = ainda não foi checado a situação do Funcionario
#ALTERADO = O nome do funcionário foi alterado  

def save_func_situacao_df():
    try:
        func_situacao_df.to_excel(func_situacao_path,index=False)
        print(func_situacao_df)
    except Exception as E:
        print(E)
        


def get_func_original_df(path:str):
    func_df = pd.read_excel(path, dtype=str).fillna("")
    return func_df

def create_func_situacao_df(path:str):
    '''Create a new dataframe with the situations of the employees'''
    
    df = pd.read_excel(path, dtype=str).fillna("")
    df["SITUACAO"] = situacao_lista[0]
    
    #Criamos a planilha na pasta data
    df.to_excel(func_situacao_path,index=False)
    
    return df
    
def get_func_situacao(path):  
    
    # Carregue os cabeçalhos da planilha existente
    df = pd.read_excel(path, dtype=str).fillna("")
        
    #Preenchemos as colunas vazias de situação por "situacao_lista[0]"
    df.loc[df['SITUACAO'] == '', 'SITUACAO'] = situacao_lista[0]    
    
    return df

# == == DF
func_situacao_df = None

if os.path.exists(func_situacao_path):
    # Carregue os cabeçalhos da planilha existente
    func_situacao_df = get_func_situacao(func_situacao_path)
       
else: 
    func_situacao_df = create_func_situacao_df(func_path)

if __name__ == '__main__':
    
    print(get_func_original_df(func_path))
    print("=="*30)
    
    func_situacao_df.loc[func_situacao_df["MATRICULA"]== "1", "NOME"]="HOJE HOJE"
    print(func_situacao_df)
    