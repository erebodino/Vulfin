import sqlite3
import os
import logging.config
import traceback
import sys
from time import sleep



logging.config.fileConfig('logger.ini', disable_existing_loggers=False)
logger = logging.getLogger(__name__)

class ManagerSQL:
    
    def __init__(self):
        self.path = os.path.join(os.getcwd(), 'RRHH_VULCANO.db')
        
        
    def conexion(self):
        logger.info('Iniciando conexion')
        conn = None
        try:
            conn = sqlite3.connect(self.path,uri=True)
            return conn
        except Exception as e:
            if sqlite3.OperationalError == type(e):
                if e.args[0].startswith('unable to open database file'):
                    logger.warning("No esta la BD")
                    print('ERROR, la base de datos ha sido comprometida, se procede al cierre')
                    sleep(5)
                    sys.exit()
                else:
                    logger.error("excepcion desconocida: %s", traceback.format_exc())
    
        return conn
    
    def executeQuery(self,conn,query):
        try:
            logger.info('Ejecutando query')
            c = conn.cursor()
            c.execute(query)
            conn.commit()
            conn.close()
        except Exception as e:
            if sqlite3.IntegrityError == type(e):              
                print('Legajo repetido, por favor cambiarlo')
                logger.warning("excepcion por legajo duplicado")
            elif sqlite3.OperationalError == type(e):
                if e.args[0].startswith('no such table'):
                    logger.warning("no se encuentran las tablas de la BD")
                    print('ERROR, la base de datos ha sido comprometida, se procede al cierre')
                    sleep(5)
                    sys.exit()
            else:
                logger.error("excepcion desconocida: %s", traceback.format_exc())