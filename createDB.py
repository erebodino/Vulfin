import sqlite3
import os
import pandas as pd
#sqlite3.connect('RRHH_VULCANO.db')
db_file = os.path.join(os.getcwd(), 'RRHH_VULCANO.db')
conn = sqlite3.connect(db_file)


# conn.execute(insertUno)
# conn.commit()

#datos = pd.read_excel(r'J:\Emma\14. Vulcano\RelojRRHH\Proyecto\BD Modif.xlsx')

class ManagerSQL:
    
    def __init__(self):
        self.path = os.path.join(os.getcwd(), 'RRHH_VULCANO.db')
        
    def conexion(self):
        conn = None
        try:
            conn = sqlite3.connect(self.path)
            return conn
        except Exception as e:
            print(e)
    
        return conn
    
    def executeQuery(self,conn,query):
        try:
            c = conn.cursor()
            c.execute(query)
            conn.commit()
            conn.close()
        except Exception as e:
            if sqlite3.IntegrityError == type(e):              
                print('Legajo repetido, por favor cambiarlo')
            else:
                print(e,type(e))

# dbObject = ManagerSQL()
# conexion = dbObject.conexion()
# dbObject.executeQuery(conexion,insertTres)
# dbObject.executeQuery(conexion,insertCuatro)
# dbObject.executeQuery(conexion,insertCinco)


# # datos.to_sql('legajos', con=conexion,if_exists='replace', index = False)
# datosDesdeSQL = pd.read_sql("SELECT * from legajos", con=conexion)

# datosRotativos = datosDesdeSQL.loc[(datosDesdeSQL['AREA']=='INYECCION') | (datosDesdeSQL['AREA']=='MECANIZADO')]
 
# datosSinRotativos = datosDesdeSQL.loc[(datosDesdeSQL['AREA']!='INYECCION') & (datosDesdeSQL['AREA']!='MECANIZADO')]