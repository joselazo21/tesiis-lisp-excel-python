import sys
from dibujar_excel import Tabla, Hoja, generate_code

if __name__ == "__main__":
    # Creacion de Tablas para hoja C111_LISP
    horario_C111_LISP = Tabla(
        id="horario_C111_LISP",
        filas=6,
        columnas=6,
        start_row=2,
        start_col=2,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Turnos", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
        nombres_filas=["Turno 1", "Turno 2", "Turno 3", "Turno 4", "Turno 5", "Turno 6"],
        contenido_de_la_tabla=None
    )
    asignaturas_C111_LISP = Tabla(
        id="asignaturas_C111_LISP",
        filas=4,
        columnas=2,
        start_row=10,
        start_col=2,
        alto_celda=None,
        ancho_celda=None,
        nombres_columnas=["Asignatura", "Frecuencia"],
        nombres_filas=None,
        contenido_de_la_tabla=[["Álgebra Lineal-C", 1], ["Álgebra Lineal-CP", 1], ["Lógica-C", 1], ["Lógica-CP", 1]]
    )
    hoja_C111_LISP = Hoja(
        grupo="C111_LISP",
        horario=horario_C111_LISP,
        asignaturas=asignaturas_C111_LISP
    )
    generate_code(hoja_C111_LISP, "Salida_C111_LISP.xlsx")

