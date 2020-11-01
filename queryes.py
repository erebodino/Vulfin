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


queryLegajo = """INSERT INTO legajos(legajo,apellido,nombre,area,tipoPago)
VALUES ('{}','{}','{}','{}','{}');"""

insertDos = """INSERT INTO legajos(legajo,apellido,nombre,area,tipoPago)
VALUES ('5','Sanchez','Elvio Eduardo','Inyeccion','Jornal');"""

insertRegistros = """INSERT INTO REGISTROS(legajo,nombre,dia,fecha,ingreso0,egreso0,ingreso1,egreso1,ingreso2,egreso2,ingreso3,egreso3,ingreso4,egreso4)
VALUES ('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}');"""

selectAll = """SELECT * from REGISTROS where fecha >= '{}' AND fecha <= '{}' order by fecha ASC"""
selectSome = """SELECT * from REGISTROS where (fecha >= '{}' AND fecha <= '{}') and (legajo in {}) order by fecha ASC"""