import pyinputplus as pyip

def validador(eleccion):
    if eleccion not in ['1','2','3']:
        raise Exception('\nEleccion de turno invalida, reingrese una de las opciones sugeridas\n\n')
def postValidacion(eleccion):
    valores = {'1':'Mañana','2':'Tarde','3':'Noche'}
    print('Su eleccion fue: ',valores[eleccion])

asd = pyip.inputCustom(customValidationFunc=validador,prompt='Ingrese el turno del operario:\n1. Mañana\n2. Tarde\n3. Noche',
                        postValidateApplyFunc=postValidacion)

