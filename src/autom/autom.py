import pyautogui as pag

class AutomException(Exception):
    def __init__(self, code:int, msg: str) -> None:
        """
        0: automação parada pelo usuario
        """
        super().__init__(f"{code}: {msg}")


def configure():
    pag.FAILSAFE = True
    pag.PAUSE = 0.6

def moveCursor(x:int,y:int, duration:float=0.0)->None:
    pag.moveTo(x,y, duration=duration)

def click(clicks_quanty:int=1)->None:
    pag.click(clicks=clicks_quanty)

def write(text: str)->None:
    pag.write(text)
    
def pressEnter(quanty_press=1)->None:
    for i in range(quanty_press):
        pag.hotkey("enter")
    
def press(key:str, quanty_press=1):
    for i in range(quanty_press): 
        pag.hotkey(key)


def alert(message:str)-> bool:
    return pag.alert(message) == "OK"

def confirm(message)-> bool:
    return pag.confirm(message) == "OK"

def prompt(message:str, allow_blank:bool = False)->str:
    response = ""

    while response =="" :
        response = pag.prompt(message)
        
        if allow_blank:
            break
    
    return "CANCEL" if response ==None else response 

def custom_message(message:str,buttons:list[str])-> bool:
    return pag.confirm(message,buttons) 
    

configure()

if __name__ == "__main__":
    
    if(confirm("Vamos começar o programa?\n Abra o seu navegador")): 
        
        site = prompt("Deseja entrar em qual site? ")
        
        if site == "CANCEL":
            raise AutomException(0,"Programa cancelado")
        
        moveCursor(265,67,2)
        click(3)
        
        write(site)
        press("down")
        pressEnter()
        
        moveCursor(363,309,2)
        click()
        alert("DIVIRTA-SE!")
        im1 = pag.screenshot()
    else:
        print("Ok :(")
    
    print("Fim do programa!")