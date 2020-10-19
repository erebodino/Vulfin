import pandas as pd
#from limpieza import ingreso_egreso
from paths import empleados_text
import pyinputplus as pyip
from analizador import Analizador
def validador(eleccion):
    if eleccion not in ['1','2','3']:
        raise Exception('\nEleccion de turno invalida, reingrese una de las opciones sugeridas\n\n')
def postValidacion(eleccion):
    valores = {'1':'Mañana','2':'Tarde','3':'Noche'}
    print('Su eleccion fue: ',valores[eleccion])
    
def ingreso_egreso(line,frame,legajo,nombre):
    try:
        fecha = pd.to_datetime(line.split()[1],yearfirst=True,format='%d/%m/%Y').date()#Da formato datetime.date
        ingresos_egresos = []
        for x in range(2,12,1):#Esta parte se encarga de dar formato al horario, siempre con dia y horas.
            try:
                ingresos_egresos.append(pd.to_datetime(('{} {}').format(fecha,line.split()[x])))            
            except:
                ingresos_egresos.append(pd.to_datetime(('{} 00:00').format(fecha)))
    

        lista_final = []
        lista_final.append(ingresos_egresos[0])
            
        for indice in range(len(ingresos_egresos) -1):#Elimina registros duplicados,usa como limite unos 5 minutos.            
            if ingresos_egresos[indice + 1] - ingresos_egresos[indice] > pd.Timedelta(minutes=5):
               lista_final.append(ingresos_egresos[indice + 1])
           
        for u in range((10-len(lista_final))):
            lista_final.append(pd.to_datetime(('{} 00:00').format(fecha)))

            # turno = pyip.inputCustom(customValidationFunc=validador,prompt='Ingrese el turno del operario:\n1. Mañana\n2. Tarde\n3. Noche',
            #             postValidateApplyFunc=postValidacion)
            # turno = input('Seleccione un turno:\n1.Mañana\n2.Noche\n')
        frame = frame.append({'Empleado':legajo,'Nombre':nombre,'Dia':line.split()[0],'Fecha':fecha,
                                  'Ingreso_0':lista_final[0],'Egreso_0':lista_final[1],
                                  'Ingreso_1':lista_final[2],'Egreso_1':lista_final[3],
                                  'Ingreso_2':lista_final[4],'Egreso_2':lista_final[5],
                                  'Ingreso_3':lista_final[6],'Egreso_3':lista_final[7],
                                  'Ingreso_4':lista_final[8],'Egreso_4':lista_final[9]},
                                  ignore_index=True) #agrega una fila mas al dataframe con los datos de ese dia
                                
        return frame
    except Exception as e:
        print(e)
        #logger.exception('',exc_info=True)



def frameFichadas():
    
    semaine = {'Lunes': 0,
                       'Martes' : 1,
                       'Miércoles': 2,
                       'Jueves':3,
                       'Viernes':4,
                       'Sábado':5,
                       'Domingo':6,
              }
    
    columnas = ['Empleado','Nombre','Dia','Fecha','Ingreso_0','Egreso_0','Ingreso_1','Egreso_1',
                        'Ingreso_2','Egreso_2','Ingreso_3','Egreso_3','Ingreso_4','Egreso_4',
                        ]
    MedioDia = None
    # fechaInicioAnalisis = pd.to_datetime(pyip.inputDate('Primer dia de analisis en formato DD/MM/AAAA: ',formats=["%d/%m/%Y"]))
    # fechaFinAnalisis = pd.to_datetime(pyip.inputDate('Ultimo dia de analisis en formato DD/MM/AAAA: ',formats=["%d/%m/%Y"]))
    frame = pd.DataFrame(columns=columnas)
    with open(empleados_text) as file: 
    #with open(empleados_text,encoding="utf-8") as file:    
                    for line in file.readlines():
                        if line.startswith('Empleado'):
                            legajo = line.split()[1].replace('.',"")
                            nombre = line.split()[3] +' '+ line.split()[4]+' '+ line.split()[5]
                        
                        for jour in semaine.keys():        
                            if line.startswith(jour):
                                frame = ingreso_egreso(line,frame,legajo,nombre)

    return frame

def logicaRotativos(frame,fechaInicio=None,fechaFin=None):
    """
    Esta funcion se encarga de limpiar los registros de los dataframes individuales de cada
    empleado rotativo, aqui dentro esta toda la logica de limpieza de esos registros.

    Parameters
    ----------
    frame : DataFrame
        Dataframe con los registros de un empleado Rotativo.

    fechaInicio : datetime.date
        Fecha de inicio para el analisis.
    fechaFin : TYPE
        DESCRIPTION.

    Returns
    -------
    frame : DataFrame
        DataFrame ya corregido. Los lugares vacios corresponden a faltas en los fichajes.

    """

    mascara = (frame['Fecha'] >= fechaInicio) & (frame['Fecha'] <= fechaFin) #mascara para filtrar el frame en funcion de la fecha de inicio y fin
    frameOriginal = frame
    frameEnAnalisis = frame.loc[mascara].copy()
    limpiador = Analizador(frameOriginal=frameOriginal,frameEnAnalisis=frameEnAnalisis)
    newFrame = limpiador.limpiador()
    
        
    return newFrame 

def frameAnalisisIndividual(frame):
    """
    Esta funcion se encarga de tomar el frame completo de todas las horas e ir armando 1 frame por cada empleado
    y si esta dentro de los rotativos encara con una logica distinta.    
    En caso de ser normal y no fichar salida o ingreso queda el espacio vacio.    
    Por ultima genera un dataframe totalizando todo
    
    Returns
    -------
    Dataframe

    """
    legajos = frame['Empleado'].unique()
    legajos_query = ['253','260','261']
    
    fechaInicio = pyip.inputDate('Ingrese el primer dia habil DD/MM/AAAA: ',formats=["%d/%m/%Y"])
    fechaFin = pyip.inputDate('Ingrese el ultimo dia habil  DD/MM/AAAA: ',formats=["%d/%m/%Y"])
    
    lista =[]
    for legajo in legajos:
        newFrame = frame[frame['Empleado']==legajo]
        if legajo in legajos_query:
            newFrame = logicaRotativos(newFrame,fechaInicio=fechaInicio,fechaFin=fechaFin)
        
        lista.append(newFrame)
    
    return lista
    

    
    
    
    
    
    
    


valor = frameFichadas()
legajos = frameAnalisisIndividual(valor)
valdez = legajos[0]
montivero = legajos[1]