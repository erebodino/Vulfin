query = """ CREATE TABLE IF NOT EXISTS legajos (
	id integer PRIMARY KEY,
    legajo integer type UNIQUE,
	apellido text NOT NULL,
    nombre text NOT NULL,
    area text NOT NULL,
    tipoPago text NOT NULL
);
"""

queryTabla = """ CREATE TABLE IF NOT EXISTS REGISTROS (
	id integer PRIMARY KEY,
    legajo integer type,
    nombre text NOT NULL,
    dia text NOT NULL,
    fecha text NOT NULL,
    ingreso0 text not NULL,
    egreso0 text NOT NULL,
    ingreso1 text not NULL,
    egreso1 text NOT NULL,
    ingreso2 text not NULL,
    egreso2 text NOT NULL,
    ingreso3 text not NULL,
    egreso3 text NOT NULL,
    ingreso4 text not NULL,
    egreso4 text NOT NULL,
    UNIQUE(legajo,fecha) ON CONFLICT IGNORE
);
"""

queryConsultaEmpleados = """SELECT * FROM LEGAJOS """

insertRegistros = """INSERT INTO REGISTROS(legajo,nombre,dia,fecha,ingreso0,egreso0,ingreso1,egreso1,ingreso2,egreso2,ingreso3,egreso3,ingreso4,egreso4)
VALUES ('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}');"""

selectAll = """SELECT * from REGISTROS where fecha >= '{}' AND fecha <= '{}' order by fecha ASC"""
selectSome = """SELECT * from REGISTROS where (fecha >= '{}' AND fecha <= '{}') and (legajo in {}) order by fecha ASC"""

insertEmpleado = """INSERT INTO legajos(LEG,APELLIDO,NOMBRE,AREA,TIPO_DE_PAGO)
VALUES ('{}','{}','{}','{}','{}');"""

deleteEmpleado = """DELETE FROM legajos
WHERE LEG ={}"""

actualizarEmpleado = """UPDATE legajos
SET {} = '{}'
WHERE LEG = '{}'
"""

# UPDATE REGISTROS
# SET 'ingreso0' = '2020-10-01 06:32:00'
# WHERE legajo = '6' and fecha = '2020-10-01'

# select * from REGISTROS
# WHERE legajo = '6' and fecha = '2020-10-01'

if __name__ == '__main__':
    pass