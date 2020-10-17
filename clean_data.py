import pandas as pd
#from limpieza import ingreso_egreso
from paths import empleados_text


def ingreso_egreso(line,frame,legajo,nombre,medioDia,rotativo=False):
    try:
        fecha = line.split()[1]
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
                
        if not rotativo:
            frame = frame.append({'Empleado':legajo,'Nombre':nombre,'Dia':line.split()[0],'Fecha':line.split()[1],
                                  'Ingreso_0':lista_final[0],'Egreso_0':lista_final[1],
                                  'Ingreso_1':lista_final[2],'Egreso_1':lista_final[3],
                                  'Ingreso_2':lista_final[4],'Egreso_2':lista_final[5],
                                  'Ingreso_3':lista_final[6],'Egreso_3':lista_final[7],
                                  'Ingreso_4':lista_final[8],'Egreso_4':lista_final[9]},
                                  ignore_index=True) #agrega una fila mas al dataframe con los datos de ese dia

        else:    
            frame = frame.append({'Empleado':legajo,'Nombre':nombre,'Dia':line.split()[0],'Fecha':line.split()[1],
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



def ingreso():
    
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
    frame = pd.DataFrame(columns=columnas)
    with open(empleados_text) as file: 
    #with open(empleados_text,encoding="utf-8") as file:    
                    for line in file.readlines():
                        if line.startswith('Empleado'):
                            legajo = line.split()[1].replace('.',"")
                            nombre = line.split()[3] +' '+ line.split()[4]+' '+ line.split()[5]
                        
                        for jour in semaine.keys():        
                            if line.startswith(jour):
                                if legajo in ['253','260','261']:
                                    frame = ingreso_egreso(line,frame,legajo,nombre,MedioDia,rotativo=True)
                                else:
                                    frame = ingreso_egreso(line,frame,legajo,nombre,MedioDia,rotativo=False)
    return frame

valor = ingreso()