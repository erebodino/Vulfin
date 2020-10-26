import pyinputplus as pyip
import pandas as pd

"""
Rrealize cambios en E:\Anaconda\envs\RelojVul\Lib\site-packages\pysimplevalidate en la funcion

validateDate, agregue un exsMsg 'No representa una fecha valida'

se modifico a nivel lenguaje

"""
def validador(eleccion):
    if eleccion not in ['1','2','3']:
        raise Exception('\nEleccion de turno invalida, reingrese una de las opciones sugeridas\n\n')

def postValidacion(eleccion):
    valores = {'1':'Mañana','2':'Tarde','3':'Noche'}
    print('Su eleccion fue: ',valores[eleccion])
    
    

# asd = pyip.inputCustom(customValidationFunc=validador,
#                        prompt='Ingrese el turno del operario:\n1. Mañana\n2. Tarde\n3. Noche',
#                        postValidateApplyFunc=postValidacion)

bcd = pd.to_datetime(pyip.inputDate('Ingrese el primer dia habil de la semana en formato DD/MM/AAAA: ',
                                    formats=["%d/%m/%Y"]))


