class Funcionario:
    matricula=""
    nome=""
    setor=""
    cpf=""
    situacao = ""
    
    def __init__(self,matricula:str,nome:str,setor:str,cpf:str,situacao:str) -> None:
        self.matricula= matricula
        self.nome= nome
        self.setor= setor
        self.cpf= cpf
        self.situacao= situacao
        
    
    def __str__(self) -> str:
        return f"MATRICULA: {self.matricula}\nNOME: {self.nome}\nSETOR: {self.setor}\nCPF: {self.cpf}\n SITUAÇÃO: {self.situacao}"