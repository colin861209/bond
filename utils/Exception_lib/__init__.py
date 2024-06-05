import os
DEBUGFLAG = False

class DateException(Exception):
    def __init__(self) -> None:
        super().__init__()

class WebException(Exception):
    def __init__(self) -> None:
        super().__init__()
        
def DEBUG(msg:str, flag=DEBUGFLAG):
    if flag: print(f'DEBUG msg: \n  {msg}\n')

def ERROR(Exception, msg:str):
    if Exception == AssertionError: print(f'ASSERT ERROR:: {msg}')
    if Exception == ValueError:     print(f'VALUE ERROR:: {msg}')
    if Exception == DateException:  print(f'DATE ERROR:: {msg}')
    if Exception == WebException:   print(f'WEB ERROR:: {msg}')
    os.system("pause")
    exit(0)

def DONE(msg:str, end=''):
    print(f'{msg}... DONE{end}')
