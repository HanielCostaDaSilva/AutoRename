import time
import autom.autom as a


matriculas =[2354567,6854697,1237849]

#time.sleep(2)


for m in matriculas:
    if not a.confirm(f"continue =>{m}"):
        break
    a.moveCursor(1124,849) #matricula
    a.click(2)
    a.write(str(m))
    a.press_enter()
    a.moveCursor(232,288) #nome
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    