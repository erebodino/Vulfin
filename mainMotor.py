import pyinputplus as pyip

def mainLoop():
    print('-'*100)
    print(' '*50,'Bienvenido',' '*50)
    print('-'*100)
    
    
    tareas = ['Ordenado de registros','Creación de informes','Gestion de Base de datos','Salir']
    tareasOrdenado = ['Limpieza de registros','Actualización de registros','Volver','Salir']
    tareasInformes = ['Ingreso de fechas','Volver','Salir']
    tareasBD = ['Actualizar','Descargar','Volver','Salir']
    
    
    respuesta = pyip.inputMenu(tareas,prompt='¿Que desea hacer?\n',lettered=True)
    if respuesta == 'Ordenado de registros':
        ordenadoRespuesta = pyip.inputMenu(tareasOrdenado,prompt='¿Que desea hacer?\n',lettered=True)
        
        
        if ordenadoRespuesta == 'Salir':
            notDone = False
        else:
            notDone = True
    
    
    elif respuesta == 'Creación de informes':
        informesRespuesta = pyip.inputMenu(tareasInformes,prompt='¿Que desea hacer?\n',lettered=True)
        if informesRespuesta == 'Salir':
            notDone = False
        else:
            notDone = True
    
    
    elif respuesta == 'Gestion de Base de datos':
        baseDeDatosRespuesta = pyip.inputMenu(tareasBD,prompt='¿Que desea hacer?\n',lettered=True)
        if baseDeDatosRespuesta == 'Salir':
            notDone = False
        else:
            notDone = True
    
    else:
        notDone = False
    return notDone

        
    
    
    
if __name__ == '__main__':
    notDone = True
    while notDone:
        notDone = mainLoop()
    SystemExit()