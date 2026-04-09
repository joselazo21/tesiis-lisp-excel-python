import sys
from dibujar_excel import Tabla, Hoja, generate_code

if __name__ == "__main__":
    # Creacion de Tablas para hoja C111
    horario_C111 = Tabla(
        id="horario_C111",
        filas=12,
        columnas=5,
        start_row=2,
        start_col=2,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
        nombres_filas=["1er Turno", "", "2do Turno", "", "3er Turno", "", "4to Turno", "", "5to Turno", "", "6to Turno", ""],
        contenido_de_la_tabla=[["Pro-C", "L-C", "A1-C", "Pro-C", ""], [7, 8, 6, 7, ""], ["A1-CP", "EF", "L-CP", "A1-C", ""], [3, "SEDER", 8, 6, ""], ["AM1-CP", "AM1-C", "Pro-CP", "Pro-CP", ""], [4, 6, 2, 2, ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "F-C", ""], ["", "", "", 9, ""], ["", "", "", "F-CP", ""], ["", "", "", "Lab", ""]]
    )
    asignaturas_C111 = Tabla(
        id="asignaturas_C111",
        filas=11,
        columnas=4,
        start_row=2,
        start_col=9,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Asignaturas", "Frec", "Faltan", "Asignadas"],
        nombres_filas=None,
        contenido_de_la_tabla=[["Pro-C", 1, 0, 1], ["Pro-CP", 2, 0, 2], ["A1-C", 2, 0, 2], ["A1-CP", 1, 0, 1], ["AM1-C", 1, 0, 1], ["AM1-CP", 1, 0, 1], ["L-C", 1, 0, 1], ["L-CP", 1, 0, 1], ["F-C", 1, 0, 1], ["F-CP", 1, 0, 1], ["EF", 1, 0, 1]]
    )
    hoja_C111 = Hoja(
        grupo="C111",
        horario=horario_C111,
        asignaturas=asignaturas_C111
    )

    # Creacion de Tablas para hoja C112
    horario_C112 = Tabla(
        id="horario_C112",
        filas=12,
        columnas=5,
        start_row=2,
        start_col=2,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
        nombres_filas=["1er Turno", "", "2do Turno", "", "3er Turno", "", "4to Turno", "", "5to Turno", "", "6to Turno", ""],
        contenido_de_la_tabla=[["L-C", "Pro-C", "AM1-CP", "F-CP", ""], [6, 2, 5, 5, ""], ["AM1-C", "F-C", "EF", "Pro-CP", "Pro-CP"], [5, 6, "SEDER", 5, "Lab"], ["A1-C", "A1-C", "L-CP", "A1-CP", ""], [8, 6, 2, 5, ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
    )
    asignaturas_C112 = Tabla(
        id="asignaturas_C112",
        filas=11,
        columnas=4,
        start_row=2,
        start_col=9,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Asignaturas", "Frec", "Faltan", "Asignadas"],
        nombres_filas=None,
        contenido_de_la_tabla=[["Pro-C", 1, 0, 1], ["Pro-CP", 2, 0, 2], ["A1-C", 2, 0, 2], ["A1-CP", 1, 0, 1], ["AM1-C", 1, 0, 1], ["AM1-CP", 1, 0, 1], ["L-C", 1, 0, 1], ["L-CP", 1, 0, 1], ["F-C", 1, 0, 1], ["F-CP", 1, 0, 1], ["EF", 1, 0, 1]]
    )
    hoja_C112 = Hoja(
        grupo="C112",
        horario=horario_C112,
        asignaturas=asignaturas_C112
    )

    # Creacion de Tablas para hoja C113
    horario_C113 = Tabla(
        id="horario_C113",
        filas=12,
        columnas=5,
        start_row=2,
        start_col=2,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
        nombres_filas=["1er Turno", "", "2do Turno", "", "3er Turno", "", "4to Turno", "", "5to Turno", "", "6to Turno", ""],
        contenido_de_la_tabla=[["L-C", "Pro-C", "AM1-CP", "F-CP", ""], ["Lab", 7, 8, 1, ""], ["AM1-C", "F-C", "EF", "Pro-CP", ""], [6, 6, "SEDER", 1, ""], ["A1-C", "A1-C", "L-CP", "A1-CP", "Pro-CP"], [6, 6, 5, 1, "SEDER"], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]]
    )
    asignaturas_C113 = Tabla(
        id="asignaturas_C113",
        filas=11,
        columnas=4,
        start_row=2,
        start_col=9,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Asignaturas", "Frec", "Faltan", "Asignadas"],
        nombres_filas=None,
        contenido_de_la_tabla=[["Pro-C", 1, 0, 1], ["Pro-CP", 2, 0, 2], ["A1-C", 2, 0, 2], ["A1-CP", 1, 0, 1], ["AM1-C", 1, 0, 1], ["AM1-CP", 1, 0, 1], ["L-C", 1, 0, 1], ["L-CP", 1, 0, 1], ["F-C", 1, 0, 1], ["F-CP", 1, 0, 1], ["EF", 1, 0, 1]]
    )
    hoja_C113 = Hoja(
        grupo="C113",
        horario=horario_C113,
        asignaturas=asignaturas_C113
    )

    hojas = [hoja_C111, hoja_C112, hoja_C113]
    generate_code(hojas, "Reporte_ODS_Final.xlsx")
