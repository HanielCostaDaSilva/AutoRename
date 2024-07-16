class Funcionario:
    matricula=""
    nome=""
    cpf=""
    situacao = ""
    
    def __init__(self,matricula:str,nome:str,cpf:str,situacao:str) -> None:
        self.matricula= matricula
        self.nome= nome
        self.cpf= cpf
        self.situacao= situacao
        
    
    def __str__(self) -> str:
        return f"MATRICULA: {self.matricula}\nNOME CORRETO: {self.nome}\nCPF: {self.cpf}\n SITUAÇÃO: {self.situacao}"