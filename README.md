# Vulfin
App de RRHH para contabilizar horas y entregar informes

Software diseñado para la gestión de horarios, por parte del área de RRHH.

Este software fue diseñado con el objetivo de procesar un txt con las horas trabajadas por operarios de turnos fijos y rotativos. Se procede a limpiar los registros corrompidos, luego se procede a ordenarlos en funcion de los turnos y las particularidades de la empresa, Y por ultimo se los registra en una base de datos para darle consistencia a los informes a lo largo del tiempo.

El sistema se ejecuta mostrando una consola de comandos en donde el usuario va seleccionando las opciones de ejecucion (requerimiento por parte de la empresa). En funcion de estas opciones de ejecucion es que se "limpian" los registros de horas, se ordenan, se imprimen informes o se gestiona la Base de datos con los operarios.

Se generan 2 informes distintos en formato pdf y 1 excel:

PDFs: 1.-Un pdf que especifica que operario no ha realizado una marcación. 2.-Un pdf que anuncia sobre retrasos, tardanzas y faltas.

Excel: 1.- Un excel con una hoja en donde estan todos los registros selecionados entre fechas y legajos. Otra hoja con todos los valores de horas normales, extras al 50% y al 100% junto con otros datos particulares de la gestion.
