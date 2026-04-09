from hoja_con_formulas import generar_excel_personalizado

sheets_cfg = []

sheets_cfg.append({
    'title': "D111",
    'data': [["Grupo ", "D111", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "F", "ICD", "AL", "AM I", "AL", "", "AL", "Álgebra Lineal", 3, 0, 3, "", "Aula 6*"], ["", "", "Aula 8", "Aula 7", "Aula 7*", "Aula 7*", "Aula 7*", "", "L", "Lógica", 2, 0, 2, "", "Aula 7"], ["", "", "", "", "", "", "", "", "IP", "Introducción a la Programación", 2, 0, 2, "", "Aula 7*"], ["", "Turno 2", "L", "AL", "EF", "L", "", "", "AM I", "Análisis Matemático I", 2, 0, 2, "", "Aula 8"], ["", "", "Aula 6*", "Aula 6*", "SEDER", "Aula 7", "", "", "ICD", "Introducción a la Ciencia de Datos", 2, 0, 2, "", "Lab"], ["", "", "", "", "", "", "", "", "F", "Filosofía", 2, 0, 2, "", ""], ["", "Turno 3", "IP", "AM I", "AM I", "IP", "", "", "EF", "Educación Física I", 2, 0, 2, "", ""], ["", "", "Aula 7", "Aula 6*", "Aula 7*", "Lab", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I10', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M10',
        'O4:O8'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M10', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O8', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 11, 'col': 12, 'value': "Total:"},
        {'row': 11, 'col': 13, 'value': "=COUNTA(I4:I10)"},
        {'row': 12, 'col': 12, 'value': "Σ Frec:"},
        {'row': 12, 'col': 13, 'value': "=SUM(K4:K10)"},
        {'row': 9, 'col': 14, 'value': "Total:"},
        {'row': 9, 'col': 15, 'value': "=COUNTA(O4:O8)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I10,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O8,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "D211",
    'data': [["Grupo ", "D211", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "MA", "Matemática y Aplicaciones", 2, 0, 2, "", "c 6"], ["", "", "", "", "", "", "", "", "Prb", "Probabilidades", 2, 0, 2, "", "c 7"], ["", "", "", "", "", "", "", "", "BD", "Bases de Datos", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "ED", "Estructura de Datos", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "VD", "Visualización de Datos", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "EP", "Economía Política", 2, 0, 2, "", ""], ["", "Turno 3", "", "", "", "", "", "", "EF", "Educación Física III", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "ED", "BD", "MA (EDO) cp 6 (con", "BD", "", "", "", "", "", "", "", "", ""], ["", "", "c 7", "c 7", "C211", "cp 7", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "VD", "EP", "EP", "Prb", "", "", "", "", "", "", "", "", ""], ["", "", "c 7", "c 7", "c 7", "cp 7", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "MA", "Prb", "EF 4:45pm a 5:35pm", "ED", "", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c 7", "", "cp 7", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I10', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M10',
        'O4:O5'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M10', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O5', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 11, 'col': 12, 'value': "Total:"},
        {'row': 11, 'col': 13, 'value': "=COUNTA(I4:I10)"},
        {'row': 12, 'col': 12, 'value': "Σ Frec:"},
        {'row': 12, 'col': 13, 'value': "=SUM(K4:K10)"},
        {'row': 6, 'col': 14, 'value': "Total:"},
        {'row': 6, 'col': 15, 'value': "=COUNTA(O4:O5)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I10,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O5,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "D311",
    'data': [["Grupo ", "D311", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "AE2", "RN", "TP 4", "", "", "", "AE2", "Análisis Estadístico II", 2, 0, 2, "", "c 2"], ["", "", "c 2", "c 2", "", "", "", "", "MDE", "Muestreo y Diseño de Experimentos", 2, 0, 2, "", "c 4"], ["", "", "", "", "", "", "", "", "RN", "Redes Neuronales", 2, 0, 2, "", "cp Lab2"], ["", "Turno 2", "MDE", "PL", "MDE", "RN", "", "", "PL", "Procesamiento del Lenguaje", 2, 0, 2, "", ""], ["", "", "c 2", "c 2", "cp 2", "cp Lab2", "", "", "PGVD", "Procesamiento de Grandes Volúmenes de Datos", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "TP", "Teoría Política", 2, 0, 2, "", ""], ["", "Turno 3", "PGVD", "TP", "AE2", "PGVD", "", "", "", "", "", "", "", "", ""], ["", "", "c 2", "c 4", "cp Lab2", "cp 7", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O6'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O6', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 7, 'col': 14, 'value': "Total:"},
        {'row': 7, 'col': 15, 'value': "=COUNTA(O4:O6)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O6,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "D411",
    'data': [["Grupo ", "D411", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "EIA", "", "SN", "", "", "IN", "Inteligencia de Negocios", 2, 0, 2, "", ""], ["", "", "CP 2", "2", "CP 2", "C4", "", "", "EIA", "Elementos de Inteligencia Artificial", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "CP", "Ciberseguridad y Privacidad", 2, 0, 2, "", ""], ["", "Turno 2", "IN", "SN", "IN", "SN", "", "", "CO2", "Curso Optativo II", 2, 0, 2, "", ""], ["", "", "2", "C4", "2", "C4", "", "", "ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "SN/DN", "Seguridad Nacional / Defensa Nacional", 2, 0, 2, "", ""], ["", "Turno 3", "ECTS 9 (con", "CO", "EIA 2", "", "", "", "", "", "", "", "", "", ""], ["", "", "C4", "2", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C111",
    'data': [["Grupo ", "C111", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "F", "A I", "AM I", "F", "", "", "A I", "Álgebra I", 3, 0, 3, "", "Aula 6"], ["", "", "Aula 6", "Aula 6*", "Aula 6*", "Aula 6", "", "", "L", "Lógica", 2, 0, 2, "", "Aula 6*"], ["", "", "", "", "", "", "", "", "P", "Programación", 3, 0, 3, "", "Lab"], ["", "Turno 2", "L", "A I", "EF", "L", "A I", "", "AM I", "Análisis Matemático I", 2, 0, 2, "", ""], ["", "", "Aula 6*", "Aula 6*", "SEDER", "Aula 6*", "Aula 6*", "", "F", "Filosofía", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "EF", "Educación Física I", 2, 0, 2, "", ""], ["", "Turno 3", "P", "AM I", "AM I", "P", "P", "", "", "", "", "", "", "", ""], ["", "", "Aula 6", "Aula 6*", "Aula 6*", "Aula 6", "Lab", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O6'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O6', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 7, 'col': 14, 'value': "Total:"},
        {'row': 7, 'col': 15, 'value': "=COUNTA(O4:O6)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O6,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C121",
    'data': [["Grupo ", "C121", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "P", "A I", "AM I", "P", "A I", "", "A I", "Álgebra I", 3, 0, 3, "", "Aula 5"], ["", "", "Aula 5", "Aula 5", "Aula 5", "Aula 5", "Aula 5", "", "L", "Lógica", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "P", "Programación", 3, 0, 3, "", ""], ["", "Turno 2", "F", "AM I", "EF", "L", "P", "", "AM I", "Análisis Matemático I", 2, 0, 2, "", ""], ["", "", "Aula 5", "Aula 5", "SEDER", "Aula 5", "Lab", "", "F", "Filosofía", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "EF", "Educación Física I", 2, 0, 2, "", ""], ["", "Turno 3", "L", "A I", "AM I", "F", "", "", "", "", "", "", "", "", ""], ["", "", "Aula 5", "Aula 5", "Aula 5", "Aula 5", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C122",
    'data': [["Grupo ", "C122", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "P", "A I", "AM I", "L", "P", "", "A I", "Álgebra I", 3, 0, 3, "", "Aula 1"], ["", "", "Aula 5", "Aula 5", "Aula 1", "Aula 1", "Lab", "", "L", "Lógica", 2, 0, 2, "", "Aula 5"], ["", "", "", "", "", "", "", "", "P", "Programación", 3, 0, 3, "", "Lab"], ["", "Turno 2", "F", "AM I", "EF", "AM I", "A I", "", "AM I", "Análisis Matemático I", 2, 0, 2, "", ""], ["", "", "Aula 5", "Aula 5", "SEDER", "Aula 1", "Aula 1", "", "F", "Filosofía", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "EF", "Educación Física I", 2, 0, 2, "", ""], ["", "Turno 3", "L", "A I", "P", "F", "", "", "", "", "", "", "", "", ""], ["", "", "Aula 5", "Aula 1", "Aula 1", "Aula 5", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O6'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O6', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 7, 'col': 14, 'value': "Total:"},
        {'row': 7, 'col': 15, 'value': "=COUNTA(O4:O6)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O6,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C211",
    'data': [["Grupo ", "C211", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "EDA", "Estructuras de Datos y Algoritmos I", 2, 0, 2, "", "c 6"], ["", "", "", "", "", "", "", "", "MD", "Matemática Discreta I", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "AC", "Arquitectura de computadoras", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "EDO", "Ecuaciones Diferenciales Ordinarias", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "MN", "Matemática Numérica", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "TP", "Teoría Política", 2, 0, 2, "", ""], ["", "Turno 3", "", "", "", "", "", "", "EF3", "Educación Física III", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "EDA I", "MD", "EDO", "MN", "AC", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c6", "cp 6", "cp 6", "lab", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "TP", "MN", "EDA I", "TP", "MD", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c 6", "cp 6", "c 6", "cp 6", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "EDO", "AC", "EF 4:45pm a 5:35pm", "", "", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c 6", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I10', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M10',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M10', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 11, 'col': 12, 'value': "Total:"},
        {'row': 11, 'col': 13, 'value': "=COUNTA(I4:I10)"},
        {'row': 12, 'col': 12, 'value': "Σ Frec:"},
        {'row': 12, 'col': 13, 'value': "=SUM(K4:K10)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I10,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C212",
    'data': [["Grupo ", "C212", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "EDA", "Estructuras de Datos y Algoritmos I", 2, 0, 2, "", "c 6"], ["", "", "", "", "", "", "", "", "MD", "Matemática Discreta I", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "AC", "Arquitectura de computadoras", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "EDO", "Ecuaciones Diferenciales Ordinarias", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "MN", "Matemática Numérica", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "TP", "Teoría Política", 2, 0, 2, "", ""], ["", "Turno 3", "", "", "", "", "", "", "EF3", "Educación Física III", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "EDA I", "MD", "EDO I", "MN", "AC", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c6", "cp 5", "cp 5", "lab", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "TP", "MN", "EDA I", "TP", "MD", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c 6", "cp 5", "c 6", "cp 5", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "EDO", "AC", "EF 4:45pm a 5:35pm", "", "", "", "", "", "", "", "", "", ""], ["", "", "c 6", "c 6", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I10', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M10',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M10', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 11, 'col': 12, 'value': "Total:"},
        {'row': 11, 'col': 13, 'value': "=COUNTA(I4:I10)"},
        {'row': 12, 'col': 12, 'value': "Σ Frec:"},
        {'row': 12, 'col': 13, 'value': "=SUM(K4:K10)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I10,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C311",
    'data': [["Grupo ", "C311", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "BD2", "Est", "BD2 cp", "MO", "", "", "RC", "Redes de Computadoras", 2, 0, 2, "", "Aula 9"], ["", "", "Aula 9", "Aula 9", "Aula 9", "Aula 9", "", "", "IS", "Ingeniería de Software", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "MO", "Modelos de Optimización", 2, 0, 2, "", ""], ["", "Turno 2", "IS c", "PD c", "IS c", "PD cp", "", "", "BD2", "Bases de Datos II", 2, 0, 2, "", ""], ["", "", "Aula 9", "Aula 9", "Aula 9", "Aula 9", "", "", "PD", "Programación Declarativa", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "Est", "Estadística", 2, 0, 2, "", ""], ["", "Turno 3", "RC", "MO", "RC cp", "Est cp", "", "", "", "", "", "", "", "", ""], ["", "", "Aula 9", "Aula 9", "Aula 9", "Aula 9", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C312",
    'data': [["Grupo ", "C312", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "BD2", "Est", "BD2 cp", "PD cp", "", "", "RC", "Redes de Computadoras", 2, 0, 2, "", "Aula 3"], ["", "", "Aula 9", "Aula 9", "Aula 3", "Aula 3", "", "", "IS", "Ingeniería de Software", 2, 0, 2, "", "Aula 9"], ["", "", "", "", "", "", "", "", "MO", "Modelos de Optimización", 2, 0, 2, "", ""], ["", "Turno 2", "IS c", "PD c", "IS c", "MO", "", "", "BD2", "Bases de Datos II", 2, 0, 2, "", ""], ["", "", "Aula 9", "Aula 9", "Aula 3", "Aula 3", "", "", "PD", "Programación Declarativa", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "Est", "Estadística", 2, 0, 2, "", ""], ["", "Turno 3", "RC", "MO", "RC cp", "Est cp", "", "", "", "", "", "", "", "", ""], ["", "", "Aula 9", "Aula 9", "Aula 3", "Aula 3", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O5'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O5', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 6, 'col': 14, 'value': "Total:"},
        {'row': 6, 'col': 15, 'value': "=COUNTA(O4:O5)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O5,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C411",
    'data': [["Grupo ", "C411", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "AM", "Aprendizaje de Máquinas", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "DAA", "Diseño y Análisis de Algoritmos", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "SD", "Sistemas Distribuidos", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "AE", "Asignatura Electiva", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "SN/DN", "Seguridad Nacional / Defensa Nacionaol", 2, 0, 2, "", ""], ["", "Turno 3", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "AM", "DAA", "DAA", "SN", "", "", "", "", "", "", "", "", ""], ["", "", "9", "9", "9", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "SD 9", "SN", "AM", "SN", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "9", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "ECTS 9", "AE", "SD", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "9", "9", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "C412",
    'data': [["Grupo ", "C412", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "AM", "Aprendizaje de Máquinas", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "DAA", "Diseño y Análisis de Algoritmos", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "SD", "Sistemas Distribuidos", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "AE", "Asignatura Electiva", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "SN/DN", "Seguridad Nacional / Defensa Nacionaol", 2, 0, 2, "", ""], ["", "Turno 3", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "AM", "DAA", "DAA", "SN", "", "", "", "", "", "", "", "", ""], ["", "", "9", "9", "9", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "SD 9", "SN", "AM", "SN", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "1", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "ECTS 9", "AE", "SD", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "9", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "M111",
    'data': [["Grupo ", "M111", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "F", "IM", "IAM", "IA", "IA", "", "IAM", "Introducción al Análisis Matemático", 3, 0, 3, "", "Aula 8"], ["", "", "Aula 8", "Aula 8", "Aula 8", "Aula 8", "Aula 8", "", "IA", "Introducción al Álgebra", 3, 0, 3, "", "PA Aula"], ["", "", "", "", "", "", "", "", "GA", "Geometría Analítica", 2, 0, 2, "", ""], ["", "Turno 2", "PA", "IAM", "EF", "PA Aula", "IAM", "", "PA", "Programación y Algoritmos", 2, 0, 2, "", ""], ["", "", "Aula 8", "Aula 8", "SEDER", "Lab", "Aula 8", "", "IM", "Introducción a la Matemática", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "F", "Filosofía", 2, 0, 2, "", ""], ["", "Turno 3", "GA", "GA", "IA", "GA 8", "", "", "EF", "Educación Física I", 2, 0, 2, "", ""], ["", "", "Aula 8", "Aula 8", "Aula 8", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I10', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M10',
        'O4:O5'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M10', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O5', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 11, 'col': 12, 'value': "Total:"},
        {'row': 11, 'col': 13, 'value': "=COUNTA(I4:I10)"},
        {'row': 12, 'col': 12, 'value': "Σ Frec:"},
        {'row': 12, 'col': 13, 'value': "=SUM(K4:K10)"},
        {'row': 6, 'col': 14, 'value': "Total:"},
        {'row': 6, 'col': 15, 'value': "=COUNTA(O4:O5)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I10,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O5,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J10',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "M211",
    'data': [["Grupo ", "M211", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "FVV", "Funciones de Varias Variables", 4, 0, 4, "", "c 3"], ["", "", "", "", "", "", "", "", "CAL", "Complementos de Álgebra Lineal", 3, 0, 3, "", ""], ["", "", "", "", "", "", "", "", "SP2", "Seminario de Problemas II", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "AE", "Asignatura Electiva I", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "EP", "Economía Política", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "EF", "Educación Física III", 2, 0, 2, "", ""], ["", "Turno 3", "CAL", "CAL", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "c 3", "c 3", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "FVV", "FVV", "FVV", "CAL", "", "", "", "", "", "", "", "", ""], ["", "", "c 3", "c 3", "cp 3", "cp 3", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "EP", "EP", "FVV", "", "", "", "", "", "", "", "", "", ""], ["", "", "c 7", "c 7", "cp 3", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "EF 4:45pm a 5:35pm", "SP", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "2", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "M311",
    'data': [["Grupo ", "M311", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "MN", "EDO", "MN", "FV", "FVC", "", "FVC", "Funciones de variable Compleja", 2, 0, 2, "", "Aula 4"], ["", "", "Aula 4", "Aula 4", "4", "C 4", "Aula 4", "", "IE", "Inferencia Estadística", 3, 0, 3, "", ""], ["", "", "", "", "", "", "", "", "EDO", "Ecuaciones Diferenciales Ordinarias", 3, 0, 3, "", ""], ["", "Turno 2", "FVC", "OM", "EDO", "IE", "EDO", "", "MN", "Matemática Numérica", 2, 0, 2, "", ""], ["", "", "Aula 4", "Aula 4", "Aula 4", "Aula 4", "Aula 4", "", "OM", "Optimización Matemática I", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "TP", "Teoría Política", 2, 0, 2, "", ""], ["", "Turno 3", "IE", "TP", "IE", "OM", "", "", "", "", "", "", "", "", ""], ["", "", "Aula 4", "Aula 4", "Aula 4", "Aula 4", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "M411",
    'data': [["Grupo ", "M411", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"], ["", "Turno 1", "", "", "", "", "", "", "MI", "Medida e Integración", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "GD", "Geometría Diferencial", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "HM", "Historia de la Matemática", 2, 0, 2, "", ""], ["", "Turno 2", "", "", "", "", "", "", "ECTS", "Estudios de Ciencia, Tecnología y Sociedad", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "AO2", "Asignatura Optativa II", 2, 0, 2, "", ""], ["", "", "", "", "", "", "", "", "AO3", "Asignatura Optativa III", 2, 0, 2, "", ""], ["", "Turno 3", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 4", "GD", "HM", "MI", "AO", "", "", "", "", "", "", "", "", ""], ["", "", "4", "4", "4", "2", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 5", "MI", "HM", "GD", "AO", "", "", "", "", "", "", "", "", ""], ["", "", "4", "4", "4", "3", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "Turno 6", "ECTS 9 (Con", "GD", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "C4", "4", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 14 for i in range(1, 16)},
    'range_styles': [
        {'range': 'I3:I9', 'style': {'bg_color': 'A9D18E'}},
        {'range': 'B4:B22', 'style': {'bg_color': 'F4CCCC'}}
    ],
    'table_ranges': [
        'B3:G3',
        'B4:B22',
        'C4:G22',
        'I3:M3',
        'I4:M9',
        'O4:O4'
    ],
    'horario_data_range': 'C4:G22',
    'table_block_sizes': [
        {'range': 'B3:G3', 'row_step': 1, 'col_step': 1},
        {'range': 'B4:B22', 'row_step': 3, 'col_step': 1},
        {'range': 'C4:G22', 'row_step': 3, 'col_step': 1},
        {'range': 'I3:M3', 'row_step': 1, 'col_step': 1},
        {'range': 'I4:M9', 'row_step': 1, 'col_step': 1},
        {'range': 'O4:O4', 'row_step': 1, 'col_step': 1}
    ],
    'merge_ranges': ["B4:B6", "B7:B9", "B10:B12", "B14:B16", "B17:B19", "B20:B22"],
    'table_borders': True,
    'border_color': '4F81BD',
    'border_style': 'medium',
    'formulas': [
        {'row': 10, 'col': 12, 'value': "Total:"},
        {'row': 10, 'col': 13, 'value': "=COUNTA(I4:I9)"},
        {'row': 11, 'col': 12, 'value': "Σ Frec:"},
        {'row': 11, 'col': 13, 'value': "=SUM(K4:K9)"},
        {'row': 5, 'col': 14, 'value': "Total:"},
        {'row': 5, 'col': 15, 'value': "=COUNTA(O4:O4)"},
        {'row': 23, 'col': 6, 'value': "Ocupados:"},
        {'row': 23, 'col': 7, 'value': "=COUNTA(C4:G22)/3"}
    ],
    'conditional_format_rules': [
        {
            'tipo': 'filas_pares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($I4:I9,{celda})=0)',
            'color': 'F4A460',
            'row_step': 3,
            'row_start_offset': 0
        },
        {
            'tipo': 'filas_impares',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", COUNTIF($O4:O4,{celda})=0)',
            'color': 'FFD700',
            'row_step': 3,
            'row_start_offset': 1
        },
        {
            'tipo': 'pares_con_siguiente',
            'rango': 'C4:G22',
            'formula': 'AND({celda}<>"", {celda_siguiente}="")',
            'color': 'FF0000',
            'row_step': 3,
            'next_offset': 1,
            'aplicar_a': 'siguiente'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            'color': '00FF00'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            'color': 'FF6B6B'
        },
        {
            'tipo': 'rango',
            'rango': 'J4:J9',
            'formula': 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            'color': 'FFA500'
        }
    ]
})

sheets_cfg.append({
    'title': "Aulas",
    'data': [["", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "1", "2", "3", "4", "5", "6", "7", "8", "9", "Lab"], ["", "Lunes", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"], ["", "1ro", "", "", "", "", "", "C111,C112", "", "", "", "C113"], ["", "2do", "", "", "", "", "C112", "C113", "", "", "", ""], ["", "3ro", "", "", "", "", "", "C111,C113", "", "C112", "", ""], ["", "4to", "", "", "", "", "", "", "", "", "", ""], ["", "5to", "", "", "", "", "", "", "", "", "", ""], ["", "6to", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", ""], ["", "Martes", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"], ["", "1ro", "", "C112", "", "", "", "C111", "C113", "", "", ""], ["", "2do", "C111", "", "", "", "", "C112,C113", "", "", "", ""], ["", "3ro", "", "", "", "", "", "C111,C112,C113", "", "", "", ""], ["", "4to", "", "", "", "", "", "", "", "", "", ""], ["", "5to", "", "", "", "", "", "", "", "", "", ""], ["", "6to", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", ""], ["", "Miércoles", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"], ["", "1ro", "", "", "", "", "C112", "C111", "", "C113", "", ""], ["", "2do", "", "", "", "", "", "", "", "", "", ""], ["", "3ro", "", "C112", "", "", "C113", "C111", "", "", "", ""], ["", "4to", "", "", "", "", "", "", "", "", "", ""], ["", "5to", "", "", "", "", "", "", "", "", "", ""], ["", "6to", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", ""], ["", "Jueves", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"], ["", "1ro", "C113", "", "", "", "C112", "C111", "", "", "", ""], ["", "2do", "C113", "", "", "", "C112", "C111", "", "", "", ""], ["", "3ro", "C113", "", "", "", "C112", "C111", "", "", "", ""], ["", "4to", "", "", "", "", "", "", "", "", "", ""], ["", "5to", "", "", "", "", "", "", "", "", "", ""], ["", "6to", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", ""], ["", "Viernes", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"], ["", "1ro", "", "", "", "", "", "", "", "", "", ""], ["", "2do", "", "", "", "", "", "", "", "", "", "C112"], ["", "3ro", "", "", "", "", "C111", "", "", "", "", ""], ["", "4to", "", "", "", "", "", "", "", "", "", ""], ["", "5to", "", "", "", "", "", "", "", "", "", ""], ["", "6to", "", "", "", "", "", "", "", "", "", ""]],
    'column_widths': {i: 12 for i in range(1, 13)},
    'range_styles': [
        {'range': 'B3:L3', 'style': {'bg_color': 'E6B8AF'}},
        {'range': 'B11:L11', 'style': {'bg_color': 'E6B8AF'}},
        {'range': 'B19:L19', 'style': {'bg_color': 'E6B8AF'}},
        {'range': 'B27:L27', 'style': {'bg_color': 'E6B8AF'}},
        {'range': 'B35:L35', 'style': {'bg_color': 'E6B8AF'}}
    ],
    'header_style': {'bold': True, 'align': "center", 'bg_color': 'E6B8AF'},
    'table_ranges': ["B3:L9", "B11:L17", "B19:L25", "B27:L33", "B35:L41"],
    'table_borders': True,
    'border_color': 'B3B3B3',
    'border_style': 'thick'
,
    'fernando_formulas': [
        {'cell': 'C4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=C$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=C$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=C$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=C$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=C$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=C$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=C$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=C$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=C$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=C$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=C$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=C$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=C$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=C$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=C$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=C$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=C$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=C$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=C$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=C$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=C$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=C$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=C$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=C$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=C$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=C$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=C$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=C$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=C$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=C$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=C$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=C$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=C$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=C$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=C$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=C$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=C$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=C$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=C$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=C$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=C$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=C$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=C$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=C$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=C$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=C$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=C$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=C$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=C$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=C$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=C$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=C$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=C$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=C$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=C$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=C$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=C$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=C$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=C$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=C$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=C$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=C$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=C$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=C$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=C$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=C$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=C$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=C$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=C$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=C$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=C$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=C$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=C$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=C$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=C$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=C$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=C$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=C$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=C$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=C$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=C$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=C$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=C$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=C$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=C$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=C$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=C$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=C$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=C$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=C$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=C$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=C$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=C$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=C$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=C$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=C$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=C$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=C$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=C$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=C$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=C$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=C$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=D$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=D$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=D$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=D$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=D$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=D$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=D$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=D$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=D$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=D$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=D$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=D$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=D$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=D$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=D$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=D$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=D$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=D$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=D$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=D$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=D$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=D$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=D$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=D$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=D$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=D$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=D$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=D$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=D$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=D$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=D$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=D$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=D$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=D$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=D$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=D$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=D$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=D$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=D$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=D$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=D$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=D$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=D$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=D$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=D$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=D$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=D$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=D$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=D$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=D$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=D$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=D$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=D$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=D$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=D$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=D$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=D$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=D$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=D$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=D$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=D$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=D$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=D$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=D$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=D$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=D$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=D$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=D$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=D$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=D$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=D$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=D$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=D$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=D$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=D$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=D$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=D$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=D$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=D$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=D$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=D$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=D$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=D$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=D$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=D$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=D$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=D$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=D$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=D$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=D$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=D$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=D$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=D$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=D$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=D$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=D$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=D$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=D$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=D$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=D$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=D$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=D$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=E$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=E$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=E$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=E$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=E$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=E$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=E$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=E$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=E$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=E$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=E$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=E$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=E$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=E$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=E$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=E$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=E$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=E$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=E$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=E$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=E$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=E$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=E$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=E$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=E$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=E$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=E$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=E$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=E$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=E$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=E$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=E$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=E$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=E$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=E$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=E$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=E$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=E$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=E$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=E$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=E$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=E$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=E$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=E$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=E$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=E$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=E$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=E$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=E$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=E$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=E$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=E$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=E$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=E$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=E$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=E$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=E$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=E$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=E$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=E$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=E$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=E$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=E$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=E$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=E$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=E$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=E$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=E$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=E$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=E$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=E$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=E$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=E$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=E$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=E$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=E$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=E$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=E$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=E$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=E$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=E$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=E$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=E$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=E$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=E$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=E$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=E$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=E$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=E$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=E$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=E$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=E$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=E$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=E$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=E$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=E$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=E$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=E$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=E$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=E$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=E$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=E$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=F$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=F$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=F$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=F$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=F$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=F$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=F$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=F$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=F$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=F$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=F$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=F$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=F$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=F$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=F$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=F$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=F$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=F$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=F$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=F$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=F$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=F$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=F$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=F$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=F$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=F$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=F$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=F$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=F$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=F$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=F$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=F$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=F$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=F$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=F$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=F$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=F$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=F$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=F$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=F$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=F$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=F$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=F$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=F$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=F$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=F$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=F$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=F$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=F$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=F$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=F$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=F$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=F$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=F$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=F$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=F$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=F$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=F$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=F$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=F$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=F$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=F$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=F$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=F$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=F$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=F$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=F$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=F$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=F$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=F$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=F$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=F$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=F$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=F$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=F$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=F$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=F$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=F$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=F$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=F$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=F$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=F$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=F$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=F$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=F$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=F$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=F$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=F$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=F$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=F$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=F$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=F$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=F$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=F$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=F$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=F$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=F$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=F$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=F$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=F$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=F$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=F$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=G$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=G$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=G$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=G$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=G$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=G$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=G$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=G$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=G$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=G$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=G$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=G$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=G$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=G$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=G$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=G$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=G$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=G$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=G$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=G$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=G$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=G$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=G$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=G$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=G$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=G$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=G$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=G$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=G$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=G$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=G$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=G$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=G$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=G$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=G$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=G$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=G$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=G$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=G$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=G$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=G$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=G$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=G$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=G$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=G$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=G$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=G$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=G$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=G$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=G$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=G$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=G$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=G$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=G$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=G$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=G$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=G$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=G$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=G$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=G$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=G$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=G$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=G$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=G$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=G$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=G$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=G$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=G$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=G$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=G$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=G$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=G$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=G$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=G$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=G$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=G$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=G$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=G$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=G$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=G$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=G$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=G$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=G$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=G$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=G$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=G$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=G$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=G$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=G$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=G$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=G$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=G$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=G$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=G$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=G$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=G$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=G$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=G$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=G$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=G$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=G$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=G$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=H$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=H$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=H$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=H$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=H$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=H$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=H$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=H$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=H$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=H$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=H$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=H$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=H$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=H$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=H$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=H$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=H$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=H$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=H$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=H$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=H$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=H$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=H$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=H$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=H$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=H$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=H$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=H$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=H$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=H$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=H$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=H$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=H$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=H$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=H$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=H$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=H$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=H$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=H$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=H$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=H$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=H$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=H$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=H$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=H$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=H$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=H$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=H$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=H$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=H$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=H$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=H$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=H$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=H$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=H$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=H$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=H$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=H$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=H$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=H$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=H$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=H$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=H$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=H$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=H$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=H$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=H$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=H$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=H$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=H$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=H$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=H$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=H$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=H$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=H$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=H$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=H$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=H$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=H$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=H$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=H$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=H$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=H$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=H$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=H$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=H$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=H$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=H$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=H$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=H$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=H$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=H$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=H$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=H$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=H$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=H$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=H$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=H$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=H$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=H$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=H$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=H$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=I$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=I$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=I$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=I$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=I$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=I$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=I$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=I$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=I$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=I$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=I$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=I$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=I$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=I$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=I$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=I$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=I$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=I$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=I$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=I$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=I$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=I$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=I$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=I$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=I$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=I$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=I$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=I$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=I$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=I$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=I$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=I$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=I$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=I$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=I$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=I$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=I$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=I$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=I$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=I$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=I$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=I$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=I$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=I$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=I$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=I$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=I$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=I$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=I$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=I$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=I$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=I$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=I$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=I$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=I$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=I$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=I$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=I$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=I$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=I$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=I$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=I$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=I$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=I$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=I$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=I$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=I$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=I$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=I$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=I$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=I$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=I$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=I$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=I$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=I$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=I$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=I$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=I$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=I$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=I$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=I$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=I$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=I$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=I$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=I$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=I$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=I$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=I$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=I$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=I$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=I$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=I$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=I$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=I$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=I$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=I$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=I$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=I$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=I$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=I$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=I$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=I$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=J$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=J$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=J$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=J$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=J$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=J$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=J$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=J$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=J$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=J$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=J$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=J$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=J$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=J$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=J$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=J$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=J$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=J$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=J$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=J$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=J$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=J$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=J$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=J$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=J$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=J$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=J$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=J$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=J$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=J$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=J$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=J$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=J$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=J$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=J$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=J$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=J$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=J$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=J$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=J$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=J$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=J$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=J$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=J$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=J$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=J$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=J$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=J$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=J$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=J$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=J$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=J$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=J$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=J$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=J$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=J$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=J$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=J$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=J$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=J$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=J$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=J$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=J$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=J$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=J$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=J$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=J$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=J$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=J$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=J$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=J$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=J$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=J$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=J$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=J$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=J$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=J$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=J$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=J$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=J$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=J$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=J$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=J$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=J$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=J$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=J$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=J$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=J$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=J$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=J$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=J$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=J$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=J$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=J$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=J$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=J$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=J$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=J$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=J$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=J$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=J$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=J$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=K$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=K$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=K$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=K$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=K$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=K$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=K$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=K$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=K$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=K$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=K$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=K$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=K$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=K$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=K$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=K$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=K$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=K$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=K$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=K$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=K$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=K$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=K$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=K$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=K$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=K$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=K$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=K$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=K$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=K$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=K$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=K$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=K$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=K$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=K$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=K$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=K$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=K$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=K$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=K$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=K$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=K$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=K$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=K$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=K$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=K$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=K$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=K$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=K$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=K$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=K$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=K$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=K$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=K$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=K$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=K$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=K$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=K$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=K$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=K$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=K$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=K$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=K$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=K$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=K$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=K$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=K$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=K$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=K$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=K$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=K$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=K$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=K$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=K$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=K$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=K$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=K$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=K$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=K$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=K$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=K$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=K$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=K$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=K$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=K$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=K$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=K$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=K$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=K$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=K$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=K$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=K$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=K$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=K$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=K$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=K$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=K$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=K$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=K$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=K$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=K$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=K$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L4', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$5=L$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$5=L$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$5=L$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$5=L$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$5=L$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$5=L$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$5=L$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$5=L$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$5=L$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$5=L$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$5=L$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$5=L$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$5=L$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$5=L$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$5=L$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$5=L$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$5=L$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L5', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$8=L$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$8=L$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$8=L$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$8=L$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$8=L$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$8=L$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$8=L$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$8=L$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$8=L$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$8=L$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$8=L$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$8=L$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$8=L$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$8=L$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$8=L$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$8=L$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$8=L$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L6', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$11=L$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$11=L$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$11=L$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$11=L$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$11=L$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$11=L$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$11=L$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$11=L$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$11=L$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$11=L$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$11=L$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$11=L$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$11=L$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$11=L$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$11=L$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$11=L$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$11=L$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L7', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$14=L$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$14=L$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$14=L$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$14=L$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$14=L$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$14=L$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$14=L$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$14=L$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$14=L$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$14=L$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$14=L$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$14=L$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$14=L$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$14=L$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$14=L$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$14=L$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$14=L$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L8', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$17=L$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$17=L$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$17=L$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$17=L$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$17=L$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$17=L$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$17=L$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$17=L$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$17=L$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$17=L$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$17=L$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$17=L$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$17=L$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$17=L$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$17=L$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$17=L$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$17=L$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L9', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$C$20=L$3,D111!$B$1 & \" \", \"\"),IF(D211!$C$20=L$3,D211!$B$1 & \" \", \"\"),IF(D311!$C$20=L$3,D311!$B$1 & \" \", \"\"),IF(D411!$C$20=L$3,D411!$B$1 & \" \", \"\"),IF(C111!$C$20=L$3,C111!$B$1 & \" \", \"\"),IF(C121!$C$20=L$3,C121!$B$1 & \" \", \"\"),IF(C122!$C$20=L$3,C122!$B$1 & \" \", \"\"),IF(C211!$C$20=L$3,C211!$B$1 & \" \", \"\"),IF(C212!$C$20=L$3,C212!$B$1 & \" \", \"\"),IF(C311!$C$20=L$3,C311!$B$1 & \" \", \"\"),IF(C312!$C$20=L$3,C312!$B$1 & \" \", \"\"),IF(C411!$C$20=L$3,C411!$B$1 & \" \", \"\"),IF(C412!$C$20=L$3,C412!$B$1 & \" \", \"\"),IF(M111!$C$20=L$3,M111!$B$1 & \" \", \"\"),IF(M211!$C$20=L$3,M211!$B$1 & \" \", \"\"),IF(M311!$C$20=L$3,M311!$B$1 & \" \", \"\"),IF(M411!$C$20=L$3,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=C$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=C$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=C$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=C$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=C$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=C$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=C$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=C$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=C$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=C$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=C$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=C$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=C$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=C$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=C$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=C$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=C$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=C$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=C$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=C$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=C$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=C$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=C$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=C$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=C$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=C$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=C$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=C$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=C$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=C$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=C$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=C$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=C$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=C$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=C$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=C$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=C$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=C$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=C$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=C$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=C$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=C$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=C$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=C$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=C$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=C$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=C$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=C$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=C$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=C$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=C$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=C$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=C$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=C$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=C$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=C$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=C$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=C$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=C$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=C$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=C$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=C$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=C$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=C$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=C$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=C$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=C$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=C$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=C$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=C$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=C$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=C$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=C$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=C$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=C$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=C$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=C$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=C$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=C$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=C$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=C$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=C$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=C$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=C$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=C$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=C$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=C$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=C$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=C$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=C$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=C$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=C$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=C$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=C$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=C$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=C$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=C$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=C$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=C$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=C$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=C$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=C$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=D$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=D$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=D$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=D$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=D$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=D$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=D$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=D$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=D$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=D$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=D$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=D$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=D$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=D$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=D$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=D$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=D$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=D$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=D$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=D$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=D$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=D$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=D$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=D$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=D$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=D$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=D$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=D$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=D$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=D$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=D$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=D$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=D$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=D$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=D$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=D$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=D$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=D$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=D$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=D$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=D$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=D$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=D$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=D$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=D$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=D$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=D$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=D$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=D$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=D$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=D$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=D$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=D$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=D$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=D$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=D$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=D$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=D$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=D$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=D$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=D$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=D$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=D$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=D$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=D$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=D$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=D$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=D$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=D$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=D$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=D$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=D$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=D$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=D$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=D$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=D$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=D$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=D$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=D$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=D$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=D$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=D$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=D$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=D$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=D$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=D$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=D$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=D$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=D$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=D$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=D$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=D$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=D$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=D$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=D$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=D$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=D$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=D$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=D$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=D$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=D$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=D$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=E$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=E$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=E$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=E$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=E$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=E$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=E$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=E$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=E$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=E$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=E$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=E$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=E$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=E$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=E$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=E$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=E$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=E$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=E$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=E$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=E$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=E$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=E$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=E$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=E$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=E$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=E$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=E$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=E$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=E$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=E$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=E$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=E$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=E$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=E$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=E$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=E$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=E$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=E$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=E$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=E$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=E$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=E$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=E$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=E$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=E$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=E$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=E$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=E$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=E$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=E$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=E$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=E$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=E$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=E$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=E$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=E$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=E$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=E$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=E$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=E$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=E$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=E$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=E$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=E$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=E$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=E$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=E$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=E$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=E$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=E$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=E$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=E$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=E$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=E$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=E$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=E$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=E$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=E$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=E$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=E$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=E$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=E$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=E$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=E$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=E$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=E$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=E$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=E$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=E$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=E$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=E$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=E$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=E$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=E$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=E$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=E$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=E$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=E$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=E$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=E$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=E$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=F$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=F$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=F$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=F$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=F$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=F$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=F$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=F$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=F$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=F$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=F$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=F$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=F$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=F$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=F$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=F$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=F$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=F$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=F$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=F$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=F$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=F$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=F$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=F$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=F$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=F$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=F$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=F$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=F$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=F$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=F$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=F$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=F$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=F$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=F$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=F$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=F$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=F$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=F$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=F$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=F$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=F$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=F$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=F$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=F$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=F$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=F$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=F$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=F$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=F$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=F$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=F$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=F$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=F$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=F$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=F$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=F$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=F$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=F$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=F$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=F$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=F$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=F$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=F$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=F$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=F$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=F$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=F$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=F$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=F$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=F$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=F$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=F$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=F$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=F$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=F$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=F$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=F$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=F$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=F$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=F$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=F$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=F$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=F$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=F$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=F$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=F$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=F$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=F$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=F$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=F$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=F$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=F$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=F$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=F$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=F$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=F$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=F$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=F$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=F$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=F$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=F$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=G$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=G$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=G$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=G$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=G$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=G$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=G$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=G$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=G$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=G$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=G$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=G$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=G$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=G$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=G$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=G$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=G$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=G$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=G$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=G$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=G$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=G$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=G$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=G$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=G$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=G$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=G$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=G$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=G$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=G$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=G$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=G$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=G$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=G$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=G$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=G$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=G$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=G$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=G$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=G$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=G$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=G$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=G$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=G$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=G$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=G$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=G$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=G$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=G$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=G$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=G$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=G$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=G$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=G$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=G$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=G$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=G$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=G$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=G$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=G$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=G$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=G$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=G$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=G$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=G$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=G$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=G$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=G$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=G$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=G$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=G$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=G$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=G$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=G$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=G$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=G$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=G$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=G$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=G$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=G$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=G$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=G$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=G$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=G$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=G$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=G$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=G$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=G$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=G$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=G$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=G$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=G$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=G$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=G$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=G$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=G$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=G$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=G$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=G$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=G$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=G$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=G$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=H$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=H$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=H$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=H$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=H$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=H$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=H$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=H$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=H$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=H$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=H$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=H$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=H$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=H$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=H$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=H$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=H$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=H$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=H$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=H$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=H$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=H$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=H$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=H$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=H$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=H$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=H$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=H$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=H$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=H$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=H$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=H$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=H$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=H$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=H$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=H$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=H$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=H$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=H$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=H$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=H$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=H$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=H$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=H$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=H$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=H$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=H$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=H$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=H$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=H$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=H$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=H$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=H$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=H$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=H$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=H$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=H$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=H$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=H$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=H$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=H$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=H$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=H$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=H$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=H$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=H$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=H$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=H$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=H$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=H$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=H$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=H$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=H$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=H$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=H$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=H$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=H$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=H$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=H$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=H$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=H$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=H$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=H$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=H$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=H$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=H$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=H$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=H$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=H$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=H$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=H$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=H$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=H$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=H$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=H$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=H$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=H$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=H$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=H$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=H$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=H$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=H$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=I$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=I$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=I$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=I$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=I$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=I$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=I$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=I$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=I$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=I$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=I$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=I$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=I$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=I$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=I$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=I$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=I$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=I$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=I$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=I$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=I$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=I$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=I$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=I$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=I$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=I$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=I$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=I$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=I$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=I$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=I$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=I$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=I$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=I$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=I$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=I$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=I$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=I$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=I$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=I$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=I$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=I$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=I$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=I$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=I$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=I$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=I$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=I$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=I$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=I$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=I$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=I$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=I$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=I$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=I$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=I$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=I$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=I$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=I$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=I$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=I$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=I$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=I$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=I$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=I$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=I$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=I$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=I$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=I$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=I$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=I$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=I$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=I$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=I$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=I$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=I$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=I$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=I$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=I$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=I$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=I$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=I$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=I$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=I$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=I$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=I$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=I$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=I$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=I$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=I$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=I$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=I$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=I$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=I$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=I$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=I$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=I$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=I$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=I$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=I$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=I$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=I$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=J$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=J$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=J$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=J$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=J$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=J$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=J$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=J$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=J$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=J$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=J$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=J$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=J$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=J$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=J$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=J$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=J$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=J$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=J$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=J$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=J$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=J$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=J$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=J$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=J$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=J$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=J$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=J$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=J$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=J$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=J$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=J$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=J$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=J$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=J$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=J$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=J$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=J$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=J$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=J$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=J$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=J$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=J$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=J$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=J$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=J$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=J$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=J$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=J$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=J$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=J$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=J$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=J$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=J$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=J$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=J$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=J$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=J$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=J$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=J$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=J$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=J$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=J$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=J$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=J$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=J$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=J$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=J$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=J$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=J$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=J$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=J$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=J$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=J$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=J$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=J$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=J$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=J$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=J$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=J$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=J$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=J$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=J$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=J$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=J$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=J$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=J$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=J$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=J$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=J$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=J$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=J$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=J$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=J$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=J$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=J$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=J$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=J$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=J$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=J$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=J$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=J$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=K$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=K$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=K$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=K$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=K$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=K$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=K$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=K$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=K$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=K$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=K$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=K$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=K$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=K$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=K$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=K$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=K$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=K$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=K$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=K$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=K$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=K$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=K$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=K$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=K$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=K$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=K$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=K$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=K$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=K$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=K$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=K$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=K$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=K$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=K$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=K$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=K$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=K$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=K$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=K$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=K$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=K$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=K$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=K$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=K$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=K$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=K$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=K$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=K$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=K$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=K$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=K$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=K$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=K$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=K$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=K$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=K$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=K$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=K$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=K$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=K$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=K$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=K$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=K$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=K$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=K$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=K$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=K$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=K$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=K$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=K$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=K$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=K$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=K$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=K$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=K$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=K$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=K$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=K$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=K$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=K$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=K$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=K$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=K$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=K$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=K$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=K$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=K$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=K$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=K$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=K$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=K$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=K$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=K$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=K$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=K$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=K$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=K$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=K$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=K$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=K$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=K$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L12', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$5=L$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$5=L$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$5=L$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$5=L$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$5=L$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$5=L$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$5=L$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$5=L$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$5=L$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$5=L$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$5=L$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$5=L$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$5=L$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$5=L$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$5=L$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$5=L$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$5=L$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L13', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$8=L$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$8=L$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$8=L$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$8=L$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$8=L$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$8=L$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$8=L$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$8=L$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$8=L$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$8=L$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$8=L$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$8=L$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$8=L$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$8=L$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$8=L$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$8=L$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$8=L$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L14', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$11=L$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$11=L$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$11=L$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$11=L$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$11=L$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$11=L$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$11=L$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$11=L$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$11=L$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$11=L$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$11=L$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$11=L$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$11=L$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$11=L$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$11=L$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$11=L$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$11=L$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L15', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$14=L$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$14=L$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$14=L$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$14=L$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$14=L$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$14=L$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$14=L$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$14=L$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$14=L$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$14=L$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$14=L$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$14=L$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$14=L$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$14=L$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$14=L$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$14=L$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$14=L$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L16', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$17=L$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$17=L$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$17=L$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$17=L$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$17=L$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$17=L$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$17=L$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$17=L$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$17=L$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$17=L$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$17=L$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$17=L$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$17=L$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$17=L$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$17=L$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$17=L$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$17=L$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L17', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$D$20=L$11,D111!$B$1 & \" \", \"\"),IF(D211!$D$20=L$11,D211!$B$1 & \" \", \"\"),IF(D311!$D$20=L$11,D311!$B$1 & \" \", \"\"),IF(D411!$D$20=L$11,D411!$B$1 & \" \", \"\"),IF(C111!$D$20=L$11,C111!$B$1 & \" \", \"\"),IF(C121!$D$20=L$11,C121!$B$1 & \" \", \"\"),IF(C122!$D$20=L$11,C122!$B$1 & \" \", \"\"),IF(C211!$D$20=L$11,C211!$B$1 & \" \", \"\"),IF(C212!$D$20=L$11,C212!$B$1 & \" \", \"\"),IF(C311!$D$20=L$11,C311!$B$1 & \" \", \"\"),IF(C312!$D$20=L$11,C312!$B$1 & \" \", \"\"),IF(C411!$D$20=L$11,C411!$B$1 & \" \", \"\"),IF(C412!$D$20=L$11,C412!$B$1 & \" \", \"\"),IF(M111!$D$20=L$11,M111!$B$1 & \" \", \"\"),IF(M211!$D$20=L$11,M211!$B$1 & \" \", \"\"),IF(M311!$D$20=L$11,M311!$B$1 & \" \", \"\"),IF(M411!$D$20=L$11,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=C$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=C$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=C$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=C$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=C$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=C$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=C$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=C$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=C$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=C$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=C$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=C$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=C$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=C$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=C$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=C$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=C$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=C$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=C$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=C$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=C$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=C$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=C$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=C$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=C$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=C$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=C$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=C$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=C$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=C$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=C$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=C$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=C$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=C$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=C$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=C$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=C$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=C$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=C$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=C$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=C$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=C$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=C$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=C$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=C$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=C$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=C$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=C$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=C$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=C$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=C$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=C$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=C$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=C$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=C$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=C$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=C$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=C$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=C$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=C$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=C$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=C$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=C$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=C$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=C$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=C$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=C$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=C$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=C$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=C$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=C$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=C$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=C$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=C$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=C$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=C$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=C$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=C$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=C$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=C$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=C$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=C$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=C$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=C$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=C$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=C$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=C$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=C$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=C$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=C$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=C$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=C$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=C$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=C$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=C$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=C$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=C$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=C$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=C$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=C$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=C$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=C$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=D$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=D$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=D$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=D$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=D$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=D$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=D$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=D$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=D$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=D$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=D$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=D$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=D$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=D$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=D$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=D$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=D$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=D$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=D$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=D$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=D$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=D$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=D$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=D$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=D$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=D$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=D$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=D$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=D$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=D$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=D$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=D$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=D$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=D$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=D$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=D$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=D$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=D$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=D$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=D$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=D$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=D$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=D$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=D$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=D$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=D$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=D$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=D$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=D$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=D$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=D$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=D$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=D$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=D$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=D$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=D$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=D$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=D$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=D$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=D$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=D$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=D$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=D$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=D$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=D$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=D$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=D$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=D$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=D$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=D$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=D$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=D$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=D$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=D$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=D$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=D$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=D$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=D$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=D$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=D$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=D$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=D$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=D$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=D$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=D$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=D$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=D$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=D$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=D$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=D$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=D$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=D$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=D$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=D$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=D$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=D$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=D$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=D$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=D$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=D$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=D$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=D$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=E$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=E$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=E$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=E$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=E$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=E$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=E$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=E$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=E$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=E$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=E$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=E$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=E$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=E$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=E$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=E$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=E$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=E$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=E$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=E$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=E$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=E$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=E$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=E$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=E$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=E$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=E$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=E$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=E$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=E$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=E$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=E$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=E$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=E$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=E$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=E$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=E$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=E$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=E$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=E$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=E$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=E$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=E$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=E$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=E$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=E$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=E$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=E$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=E$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=E$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=E$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=E$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=E$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=E$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=E$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=E$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=E$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=E$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=E$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=E$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=E$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=E$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=E$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=E$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=E$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=E$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=E$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=E$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=E$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=E$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=E$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=E$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=E$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=E$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=E$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=E$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=E$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=E$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=E$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=E$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=E$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=E$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=E$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=E$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=E$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=E$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=E$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=E$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=E$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=E$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=E$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=E$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=E$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=E$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=E$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=E$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=E$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=E$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=E$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=E$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=E$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=E$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=F$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=F$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=F$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=F$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=F$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=F$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=F$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=F$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=F$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=F$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=F$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=F$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=F$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=F$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=F$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=F$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=F$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=F$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=F$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=F$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=F$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=F$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=F$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=F$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=F$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=F$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=F$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=F$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=F$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=F$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=F$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=F$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=F$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=F$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=F$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=F$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=F$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=F$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=F$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=F$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=F$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=F$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=F$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=F$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=F$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=F$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=F$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=F$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=F$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=F$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=F$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=F$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=F$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=F$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=F$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=F$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=F$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=F$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=F$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=F$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=F$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=F$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=F$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=F$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=F$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=F$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=F$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=F$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=F$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=F$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=F$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=F$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=F$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=F$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=F$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=F$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=F$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=F$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=F$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=F$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=F$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=F$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=F$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=F$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=F$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=F$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=F$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=F$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=F$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=F$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=F$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=F$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=F$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=F$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=F$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=F$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=F$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=F$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=F$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=F$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=F$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=F$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=G$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=G$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=G$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=G$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=G$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=G$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=G$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=G$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=G$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=G$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=G$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=G$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=G$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=G$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=G$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=G$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=G$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=G$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=G$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=G$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=G$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=G$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=G$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=G$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=G$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=G$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=G$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=G$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=G$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=G$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=G$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=G$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=G$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=G$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=G$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=G$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=G$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=G$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=G$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=G$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=G$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=G$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=G$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=G$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=G$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=G$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=G$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=G$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=G$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=G$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=G$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=G$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=G$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=G$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=G$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=G$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=G$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=G$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=G$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=G$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=G$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=G$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=G$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=G$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=G$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=G$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=G$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=G$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=G$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=G$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=G$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=G$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=G$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=G$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=G$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=G$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=G$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=G$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=G$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=G$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=G$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=G$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=G$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=G$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=G$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=G$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=G$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=G$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=G$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=G$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=G$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=G$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=G$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=G$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=G$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=G$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=G$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=G$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=G$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=G$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=G$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=G$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=H$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=H$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=H$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=H$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=H$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=H$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=H$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=H$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=H$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=H$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=H$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=H$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=H$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=H$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=H$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=H$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=H$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=H$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=H$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=H$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=H$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=H$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=H$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=H$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=H$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=H$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=H$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=H$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=H$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=H$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=H$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=H$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=H$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=H$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=H$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=H$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=H$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=H$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=H$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=H$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=H$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=H$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=H$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=H$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=H$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=H$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=H$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=H$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=H$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=H$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=H$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=H$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=H$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=H$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=H$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=H$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=H$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=H$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=H$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=H$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=H$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=H$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=H$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=H$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=H$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=H$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=H$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=H$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=H$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=H$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=H$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=H$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=H$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=H$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=H$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=H$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=H$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=H$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=H$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=H$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=H$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=H$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=H$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=H$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=H$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=H$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=H$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=H$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=H$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=H$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=H$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=H$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=H$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=H$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=H$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=H$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=H$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=H$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=H$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=H$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=H$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=H$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=I$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=I$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=I$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=I$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=I$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=I$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=I$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=I$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=I$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=I$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=I$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=I$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=I$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=I$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=I$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=I$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=I$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=I$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=I$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=I$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=I$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=I$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=I$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=I$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=I$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=I$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=I$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=I$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=I$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=I$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=I$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=I$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=I$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=I$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=I$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=I$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=I$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=I$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=I$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=I$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=I$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=I$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=I$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=I$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=I$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=I$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=I$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=I$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=I$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=I$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=I$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=I$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=I$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=I$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=I$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=I$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=I$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=I$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=I$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=I$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=I$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=I$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=I$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=I$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=I$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=I$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=I$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=I$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=I$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=I$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=I$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=I$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=I$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=I$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=I$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=I$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=I$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=I$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=I$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=I$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=I$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=I$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=I$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=I$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=I$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=I$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=I$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=I$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=I$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=I$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=I$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=I$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=I$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=I$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=I$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=I$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=I$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=I$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=I$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=I$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=I$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=I$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=J$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=J$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=J$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=J$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=J$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=J$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=J$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=J$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=J$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=J$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=J$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=J$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=J$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=J$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=J$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=J$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=J$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=J$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=J$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=J$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=J$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=J$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=J$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=J$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=J$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=J$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=J$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=J$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=J$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=J$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=J$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=J$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=J$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=J$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=J$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=J$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=J$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=J$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=J$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=J$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=J$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=J$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=J$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=J$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=J$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=J$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=J$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=J$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=J$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=J$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=J$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=J$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=J$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=J$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=J$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=J$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=J$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=J$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=J$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=J$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=J$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=J$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=J$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=J$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=J$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=J$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=J$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=J$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=J$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=J$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=J$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=J$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=J$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=J$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=J$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=J$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=J$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=J$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=J$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=J$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=J$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=J$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=J$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=J$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=J$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=J$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=J$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=J$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=J$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=J$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=J$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=J$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=J$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=J$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=J$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=J$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=J$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=J$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=J$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=J$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=J$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=J$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=K$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=K$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=K$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=K$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=K$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=K$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=K$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=K$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=K$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=K$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=K$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=K$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=K$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=K$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=K$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=K$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=K$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=K$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=K$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=K$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=K$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=K$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=K$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=K$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=K$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=K$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=K$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=K$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=K$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=K$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=K$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=K$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=K$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=K$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=K$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=K$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=K$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=K$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=K$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=K$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=K$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=K$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=K$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=K$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=K$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=K$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=K$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=K$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=K$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=K$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=K$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=K$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=K$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=K$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=K$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=K$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=K$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=K$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=K$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=K$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=K$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=K$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=K$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=K$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=K$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=K$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=K$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=K$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=K$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=K$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=K$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=K$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=K$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=K$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=K$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=K$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=K$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=K$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=K$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=K$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=K$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=K$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=K$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=K$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=K$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=K$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=K$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=K$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=K$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=K$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=K$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=K$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=K$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=K$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=K$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=K$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=K$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=K$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=K$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=K$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=K$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=K$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L20', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$5=L$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$5=L$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$5=L$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$5=L$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$5=L$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$5=L$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$5=L$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$5=L$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$5=L$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$5=L$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$5=L$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$5=L$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$5=L$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$5=L$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$5=L$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$5=L$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$5=L$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L21', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$8=L$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$8=L$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$8=L$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$8=L$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$8=L$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$8=L$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$8=L$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$8=L$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$8=L$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$8=L$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$8=L$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$8=L$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$8=L$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$8=L$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$8=L$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$8=L$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$8=L$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L22', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$11=L$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$11=L$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$11=L$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$11=L$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$11=L$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$11=L$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$11=L$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$11=L$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$11=L$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$11=L$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$11=L$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$11=L$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$11=L$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$11=L$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$11=L$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$11=L$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$11=L$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L23', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$14=L$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$14=L$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$14=L$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$14=L$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$14=L$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$14=L$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$14=L$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$14=L$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$14=L$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$14=L$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$14=L$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$14=L$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$14=L$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$14=L$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$14=L$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$14=L$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$14=L$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L24', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$17=L$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$17=L$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$17=L$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$17=L$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$17=L$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$17=L$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$17=L$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$17=L$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$17=L$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$17=L$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$17=L$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$17=L$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$17=L$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$17=L$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$17=L$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$17=L$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$17=L$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L25', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$E$20=L$19,D111!$B$1 & \" \", \"\"),IF(D211!$E$20=L$19,D211!$B$1 & \" \", \"\"),IF(D311!$E$20=L$19,D311!$B$1 & \" \", \"\"),IF(D411!$E$20=L$19,D411!$B$1 & \" \", \"\"),IF(C111!$E$20=L$19,C111!$B$1 & \" \", \"\"),IF(C121!$E$20=L$19,C121!$B$1 & \" \", \"\"),IF(C122!$E$20=L$19,C122!$B$1 & \" \", \"\"),IF(C211!$E$20=L$19,C211!$B$1 & \" \", \"\"),IF(C212!$E$20=L$19,C212!$B$1 & \" \", \"\"),IF(C311!$E$20=L$19,C311!$B$1 & \" \", \"\"),IF(C312!$E$20=L$19,C312!$B$1 & \" \", \"\"),IF(C411!$E$20=L$19,C411!$B$1 & \" \", \"\"),IF(C412!$E$20=L$19,C412!$B$1 & \" \", \"\"),IF(M111!$E$20=L$19,M111!$B$1 & \" \", \"\"),IF(M211!$E$20=L$19,M211!$B$1 & \" \", \"\"),IF(M311!$E$20=L$19,M311!$B$1 & \" \", \"\"),IF(M411!$E$20=L$19,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=C$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=C$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=C$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=C$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=C$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=C$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=C$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=C$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=C$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=C$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=C$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=C$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=C$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=C$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=C$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=C$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=C$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=C$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=C$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=C$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=C$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=C$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=C$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=C$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=C$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=C$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=C$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=C$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=C$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=C$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=C$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=C$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=C$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=C$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=C$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=C$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=C$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=C$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=C$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=C$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=C$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=C$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=C$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=C$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=C$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=C$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=C$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=C$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=C$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=C$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=C$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=C$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=C$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=C$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=C$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=C$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=C$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=C$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=C$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=C$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=C$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=C$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=C$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=C$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=C$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=C$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=C$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=C$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=C$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=C$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=C$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=C$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=C$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=C$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=C$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=C$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=C$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=C$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=C$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=C$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=C$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=C$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=C$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=C$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=C$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=C$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=C$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=C$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=C$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=C$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=C$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=C$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=C$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=C$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=C$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=C$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=C$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=C$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=C$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=C$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=C$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=C$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=D$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=D$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=D$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=D$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=D$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=D$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=D$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=D$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=D$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=D$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=D$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=D$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=D$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=D$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=D$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=D$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=D$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=D$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=D$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=D$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=D$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=D$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=D$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=D$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=D$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=D$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=D$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=D$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=D$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=D$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=D$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=D$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=D$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=D$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=D$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=D$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=D$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=D$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=D$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=D$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=D$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=D$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=D$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=D$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=D$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=D$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=D$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=D$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=D$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=D$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=D$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=D$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=D$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=D$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=D$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=D$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=D$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=D$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=D$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=D$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=D$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=D$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=D$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=D$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=D$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=D$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=D$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=D$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=D$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=D$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=D$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=D$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=D$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=D$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=D$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=D$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=D$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=D$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=D$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=D$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=D$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=D$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=D$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=D$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=D$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=D$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=D$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=D$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=D$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=D$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=D$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=D$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=D$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=D$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=D$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=D$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=D$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=D$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=D$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=D$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=D$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=D$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=E$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=E$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=E$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=E$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=E$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=E$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=E$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=E$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=E$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=E$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=E$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=E$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=E$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=E$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=E$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=E$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=E$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=E$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=E$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=E$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=E$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=E$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=E$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=E$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=E$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=E$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=E$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=E$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=E$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=E$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=E$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=E$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=E$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=E$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=E$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=E$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=E$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=E$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=E$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=E$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=E$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=E$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=E$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=E$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=E$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=E$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=E$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=E$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=E$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=E$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=E$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=E$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=E$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=E$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=E$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=E$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=E$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=E$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=E$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=E$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=E$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=E$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=E$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=E$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=E$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=E$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=E$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=E$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=E$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=E$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=E$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=E$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=E$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=E$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=E$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=E$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=E$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=E$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=E$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=E$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=E$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=E$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=E$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=E$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=E$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=E$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=E$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=E$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=E$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=E$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=E$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=E$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=E$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=E$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=E$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=E$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=E$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=E$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=E$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=E$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=E$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=E$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=F$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=F$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=F$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=F$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=F$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=F$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=F$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=F$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=F$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=F$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=F$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=F$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=F$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=F$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=F$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=F$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=F$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=F$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=F$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=F$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=F$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=F$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=F$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=F$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=F$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=F$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=F$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=F$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=F$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=F$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=F$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=F$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=F$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=F$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=F$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=F$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=F$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=F$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=F$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=F$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=F$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=F$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=F$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=F$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=F$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=F$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=F$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=F$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=F$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=F$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=F$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=F$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=F$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=F$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=F$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=F$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=F$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=F$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=F$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=F$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=F$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=F$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=F$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=F$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=F$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=F$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=F$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=F$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=F$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=F$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=F$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=F$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=F$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=F$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=F$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=F$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=F$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=F$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=F$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=F$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=F$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=F$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=F$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=F$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=F$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=F$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=F$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=F$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=F$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=F$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=F$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=F$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=F$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=F$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=F$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=F$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=F$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=F$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=F$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=F$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=F$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=F$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=G$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=G$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=G$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=G$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=G$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=G$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=G$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=G$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=G$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=G$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=G$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=G$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=G$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=G$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=G$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=G$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=G$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=G$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=G$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=G$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=G$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=G$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=G$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=G$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=G$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=G$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=G$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=G$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=G$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=G$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=G$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=G$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=G$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=G$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=G$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=G$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=G$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=G$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=G$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=G$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=G$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=G$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=G$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=G$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=G$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=G$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=G$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=G$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=G$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=G$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=G$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=G$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=G$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=G$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=G$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=G$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=G$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=G$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=G$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=G$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=G$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=G$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=G$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=G$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=G$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=G$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=G$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=G$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=G$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=G$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=G$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=G$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=G$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=G$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=G$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=G$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=G$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=G$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=G$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=G$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=G$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=G$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=G$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=G$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=G$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=G$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=G$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=G$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=G$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=G$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=G$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=G$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=G$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=G$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=G$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=G$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=G$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=G$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=G$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=G$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=G$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=G$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=H$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=H$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=H$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=H$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=H$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=H$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=H$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=H$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=H$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=H$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=H$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=H$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=H$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=H$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=H$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=H$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=H$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=H$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=H$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=H$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=H$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=H$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=H$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=H$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=H$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=H$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=H$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=H$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=H$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=H$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=H$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=H$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=H$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=H$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=H$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=H$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=H$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=H$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=H$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=H$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=H$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=H$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=H$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=H$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=H$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=H$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=H$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=H$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=H$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=H$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=H$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=H$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=H$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=H$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=H$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=H$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=H$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=H$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=H$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=H$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=H$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=H$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=H$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=H$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=H$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=H$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=H$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=H$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=H$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=H$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=H$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=H$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=H$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=H$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=H$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=H$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=H$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=H$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=H$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=H$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=H$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=H$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=H$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=H$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=H$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=H$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=H$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=H$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=H$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=H$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=H$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=H$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=H$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=H$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=H$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=H$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=H$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=H$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=H$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=H$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=H$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=H$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=I$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=I$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=I$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=I$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=I$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=I$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=I$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=I$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=I$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=I$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=I$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=I$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=I$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=I$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=I$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=I$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=I$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=I$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=I$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=I$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=I$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=I$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=I$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=I$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=I$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=I$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=I$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=I$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=I$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=I$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=I$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=I$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=I$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=I$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=I$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=I$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=I$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=I$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=I$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=I$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=I$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=I$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=I$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=I$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=I$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=I$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=I$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=I$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=I$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=I$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=I$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=I$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=I$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=I$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=I$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=I$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=I$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=I$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=I$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=I$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=I$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=I$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=I$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=I$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=I$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=I$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=I$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=I$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=I$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=I$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=I$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=I$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=I$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=I$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=I$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=I$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=I$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=I$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=I$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=I$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=I$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=I$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=I$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=I$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=I$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=I$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=I$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=I$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=I$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=I$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=I$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=I$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=I$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=I$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=I$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=I$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=I$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=I$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=I$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=I$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=I$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=I$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=J$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=J$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=J$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=J$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=J$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=J$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=J$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=J$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=J$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=J$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=J$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=J$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=J$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=J$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=J$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=J$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=J$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=J$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=J$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=J$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=J$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=J$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=J$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=J$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=J$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=J$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=J$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=J$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=J$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=J$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=J$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=J$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=J$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=J$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=J$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=J$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=J$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=J$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=J$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=J$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=J$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=J$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=J$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=J$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=J$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=J$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=J$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=J$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=J$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=J$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=J$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=J$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=J$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=J$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=J$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=J$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=J$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=J$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=J$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=J$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=J$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=J$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=J$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=J$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=J$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=J$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=J$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=J$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=J$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=J$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=J$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=J$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=J$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=J$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=J$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=J$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=J$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=J$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=J$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=J$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=J$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=J$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=J$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=J$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=J$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=J$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=J$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=J$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=J$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=J$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=J$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=J$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=J$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=J$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=J$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=J$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=J$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=J$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=J$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=J$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=J$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=J$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=K$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=K$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=K$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=K$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=K$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=K$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=K$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=K$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=K$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=K$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=K$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=K$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=K$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=K$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=K$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=K$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=K$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=K$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=K$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=K$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=K$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=K$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=K$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=K$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=K$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=K$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=K$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=K$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=K$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=K$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=K$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=K$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=K$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=K$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=K$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=K$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=K$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=K$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=K$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=K$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=K$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=K$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=K$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=K$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=K$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=K$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=K$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=K$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=K$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=K$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=K$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=K$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=K$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=K$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=K$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=K$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=K$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=K$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=K$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=K$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=K$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=K$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=K$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=K$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=K$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=K$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=K$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=K$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=K$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=K$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=K$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=K$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=K$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=K$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=K$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=K$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=K$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=K$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=K$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=K$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=K$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=K$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=K$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=K$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=K$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=K$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=K$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=K$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=K$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=K$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=K$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=K$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=K$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=K$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=K$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=K$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=K$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=K$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=K$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=K$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=K$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=K$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L28', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$5=L$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$5=L$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$5=L$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$5=L$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$5=L$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$5=L$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$5=L$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$5=L$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$5=L$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$5=L$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$5=L$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$5=L$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$5=L$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$5=L$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$5=L$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$5=L$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$5=L$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L29', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$8=L$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$8=L$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$8=L$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$8=L$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$8=L$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$8=L$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$8=L$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$8=L$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$8=L$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$8=L$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$8=L$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$8=L$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$8=L$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$8=L$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$8=L$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$8=L$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$8=L$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L30', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$11=L$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$11=L$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$11=L$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$11=L$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$11=L$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$11=L$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$11=L$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$11=L$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$11=L$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$11=L$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$11=L$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$11=L$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$11=L$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$11=L$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$11=L$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$11=L$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$11=L$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L31', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$14=L$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$14=L$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$14=L$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$14=L$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$14=L$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$14=L$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$14=L$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$14=L$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$14=L$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$14=L$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$14=L$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$14=L$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$14=L$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$14=L$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$14=L$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$14=L$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$14=L$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L32', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$17=L$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$17=L$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$17=L$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$17=L$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$17=L$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$17=L$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$17=L$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$17=L$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$17=L$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$17=L$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$17=L$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$17=L$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$17=L$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$17=L$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$17=L$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$17=L$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$17=L$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L33', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$F$20=L$27,D111!$B$1 & \" \", \"\"),IF(D211!$F$20=L$27,D211!$B$1 & \" \", \"\"),IF(D311!$F$20=L$27,D311!$B$1 & \" \", \"\"),IF(D411!$F$20=L$27,D411!$B$1 & \" \", \"\"),IF(C111!$F$20=L$27,C111!$B$1 & \" \", \"\"),IF(C121!$F$20=L$27,C121!$B$1 & \" \", \"\"),IF(C122!$F$20=L$27,C122!$B$1 & \" \", \"\"),IF(C211!$F$20=L$27,C211!$B$1 & \" \", \"\"),IF(C212!$F$20=L$27,C212!$B$1 & \" \", \"\"),IF(C311!$F$20=L$27,C311!$B$1 & \" \", \"\"),IF(C312!$F$20=L$27,C312!$B$1 & \" \", \"\"),IF(C411!$F$20=L$27,C411!$B$1 & \" \", \"\"),IF(C412!$F$20=L$27,C412!$B$1 & \" \", \"\"),IF(M111!$F$20=L$27,M111!$B$1 & \" \", \"\"),IF(M211!$F$20=L$27,M211!$B$1 & \" \", \"\"),IF(M311!$F$20=L$27,M311!$B$1 & \" \", \"\"),IF(M411!$F$20=L$27,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=C$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=C$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=C$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=C$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=C$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=C$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=C$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=C$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=C$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=C$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=C$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=C$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=C$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=C$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=C$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=C$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=C$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=C$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=C$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=C$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=C$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=C$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=C$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=C$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=C$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=C$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=C$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=C$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=C$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=C$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=C$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=C$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=C$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=C$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=C$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=C$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=C$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=C$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=C$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=C$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=C$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=C$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=C$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=C$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=C$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=C$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=C$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=C$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=C$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=C$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=C$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=C$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=C$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=C$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=C$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=C$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=C$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=C$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=C$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=C$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=C$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=C$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=C$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=C$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=C$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=C$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=C$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=C$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=C$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=C$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=C$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=C$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=C$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=C$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=C$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=C$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=C$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=C$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=C$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=C$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=C$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=C$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=C$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=C$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=C$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'C41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=C$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=C$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=C$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=C$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=C$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=C$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=C$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=C$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=C$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=C$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=C$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=C$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=C$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=C$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=C$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=C$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=C$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=D$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=D$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=D$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=D$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=D$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=D$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=D$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=D$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=D$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=D$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=D$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=D$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=D$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=D$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=D$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=D$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=D$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=D$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=D$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=D$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=D$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=D$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=D$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=D$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=D$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=D$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=D$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=D$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=D$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=D$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=D$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=D$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=D$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=D$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=D$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=D$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=D$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=D$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=D$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=D$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=D$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=D$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=D$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=D$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=D$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=D$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=D$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=D$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=D$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=D$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=D$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=D$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=D$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=D$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=D$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=D$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=D$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=D$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=D$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=D$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=D$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=D$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=D$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=D$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=D$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=D$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=D$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=D$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=D$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=D$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=D$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=D$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=D$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=D$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=D$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=D$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=D$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=D$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=D$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=D$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=D$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=D$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=D$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=D$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=D$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'D41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=D$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=D$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=D$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=D$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=D$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=D$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=D$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=D$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=D$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=D$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=D$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=D$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=D$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=D$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=D$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=D$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=D$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=E$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=E$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=E$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=E$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=E$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=E$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=E$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=E$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=E$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=E$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=E$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=E$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=E$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=E$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=E$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=E$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=E$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=E$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=E$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=E$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=E$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=E$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=E$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=E$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=E$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=E$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=E$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=E$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=E$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=E$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=E$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=E$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=E$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=E$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=E$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=E$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=E$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=E$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=E$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=E$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=E$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=E$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=E$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=E$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=E$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=E$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=E$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=E$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=E$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=E$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=E$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=E$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=E$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=E$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=E$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=E$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=E$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=E$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=E$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=E$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=E$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=E$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=E$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=E$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=E$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=E$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=E$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=E$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=E$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=E$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=E$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=E$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=E$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=E$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=E$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=E$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=E$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=E$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=E$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=E$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=E$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=E$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=E$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=E$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=E$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'E41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=E$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=E$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=E$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=E$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=E$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=E$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=E$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=E$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=E$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=E$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=E$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=E$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=E$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=E$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=E$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=E$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=E$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=F$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=F$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=F$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=F$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=F$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=F$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=F$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=F$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=F$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=F$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=F$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=F$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=F$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=F$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=F$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=F$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=F$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=F$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=F$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=F$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=F$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=F$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=F$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=F$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=F$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=F$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=F$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=F$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=F$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=F$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=F$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=F$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=F$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=F$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=F$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=F$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=F$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=F$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=F$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=F$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=F$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=F$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=F$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=F$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=F$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=F$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=F$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=F$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=F$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=F$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=F$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=F$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=F$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=F$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=F$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=F$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=F$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=F$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=F$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=F$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=F$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=F$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=F$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=F$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=F$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=F$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=F$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=F$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=F$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=F$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=F$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=F$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=F$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=F$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=F$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=F$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=F$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=F$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=F$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=F$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=F$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=F$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=F$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=F$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=F$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'F41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=F$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=F$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=F$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=F$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=F$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=F$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=F$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=F$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=F$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=F$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=F$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=F$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=F$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=F$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=F$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=F$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=F$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=G$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=G$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=G$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=G$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=G$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=G$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=G$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=G$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=G$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=G$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=G$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=G$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=G$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=G$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=G$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=G$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=G$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=G$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=G$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=G$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=G$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=G$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=G$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=G$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=G$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=G$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=G$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=G$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=G$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=G$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=G$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=G$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=G$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=G$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=G$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=G$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=G$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=G$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=G$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=G$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=G$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=G$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=G$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=G$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=G$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=G$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=G$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=G$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=G$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=G$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=G$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=G$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=G$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=G$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=G$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=G$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=G$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=G$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=G$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=G$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=G$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=G$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=G$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=G$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=G$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=G$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=G$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=G$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=G$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=G$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=G$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=G$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=G$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=G$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=G$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=G$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=G$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=G$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=G$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=G$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=G$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=G$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=G$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=G$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=G$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'G41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=G$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=G$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=G$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=G$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=G$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=G$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=G$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=G$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=G$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=G$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=G$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=G$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=G$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=G$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=G$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=G$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=G$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=H$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=H$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=H$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=H$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=H$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=H$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=H$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=H$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=H$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=H$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=H$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=H$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=H$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=H$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=H$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=H$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=H$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=H$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=H$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=H$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=H$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=H$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=H$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=H$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=H$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=H$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=H$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=H$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=H$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=H$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=H$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=H$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=H$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=H$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=H$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=H$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=H$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=H$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=H$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=H$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=H$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=H$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=H$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=H$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=H$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=H$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=H$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=H$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=H$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=H$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=H$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=H$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=H$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=H$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=H$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=H$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=H$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=H$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=H$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=H$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=H$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=H$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=H$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=H$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=H$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=H$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=H$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=H$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=H$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=H$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=H$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=H$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=H$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=H$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=H$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=H$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=H$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=H$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=H$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=H$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=H$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=H$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=H$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=H$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=H$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'H41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=H$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=H$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=H$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=H$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=H$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=H$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=H$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=H$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=H$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=H$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=H$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=H$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=H$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=H$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=H$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=H$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=H$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=I$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=I$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=I$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=I$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=I$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=I$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=I$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=I$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=I$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=I$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=I$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=I$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=I$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=I$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=I$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=I$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=I$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=I$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=I$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=I$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=I$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=I$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=I$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=I$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=I$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=I$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=I$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=I$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=I$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=I$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=I$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=I$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=I$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=I$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=I$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=I$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=I$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=I$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=I$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=I$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=I$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=I$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=I$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=I$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=I$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=I$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=I$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=I$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=I$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=I$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=I$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=I$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=I$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=I$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=I$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=I$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=I$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=I$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=I$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=I$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=I$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=I$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=I$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=I$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=I$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=I$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=I$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=I$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=I$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=I$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=I$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=I$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=I$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=I$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=I$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=I$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=I$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=I$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=I$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=I$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=I$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=I$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=I$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=I$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=I$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'I41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=I$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=I$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=I$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=I$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=I$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=I$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=I$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=I$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=I$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=I$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=I$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=I$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=I$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=I$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=I$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=I$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=I$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=J$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=J$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=J$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=J$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=J$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=J$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=J$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=J$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=J$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=J$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=J$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=J$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=J$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=J$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=J$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=J$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=J$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=J$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=J$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=J$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=J$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=J$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=J$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=J$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=J$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=J$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=J$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=J$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=J$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=J$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=J$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=J$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=J$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=J$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=J$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=J$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=J$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=J$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=J$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=J$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=J$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=J$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=J$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=J$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=J$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=J$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=J$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=J$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=J$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=J$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=J$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=J$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=J$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=J$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=J$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=J$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=J$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=J$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=J$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=J$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=J$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=J$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=J$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=J$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=J$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=J$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=J$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=J$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=J$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=J$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=J$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=J$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=J$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=J$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=J$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=J$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=J$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=J$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=J$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=J$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=J$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=J$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=J$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=J$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=J$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'J41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=J$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=J$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=J$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=J$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=J$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=J$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=J$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=J$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=J$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=J$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=J$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=J$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=J$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=J$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=J$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=J$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=J$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=K$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=K$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=K$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=K$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=K$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=K$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=K$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=K$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=K$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=K$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=K$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=K$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=K$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=K$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=K$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=K$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=K$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=K$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=K$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=K$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=K$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=K$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=K$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=K$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=K$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=K$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=K$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=K$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=K$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=K$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=K$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=K$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=K$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=K$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=K$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=K$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=K$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=K$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=K$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=K$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=K$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=K$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=K$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=K$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=K$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=K$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=K$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=K$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=K$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=K$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=K$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=K$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=K$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=K$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=K$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=K$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=K$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=K$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=K$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=K$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=K$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=K$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=K$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=K$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=K$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=K$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=K$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=K$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=K$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=K$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=K$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=K$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=K$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=K$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=K$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=K$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=K$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=K$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=K$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=K$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=K$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=K$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=K$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=K$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=K$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'K41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=K$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=K$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=K$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=K$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=K$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=K$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=K$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=K$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=K$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=K$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=K$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=K$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=K$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=K$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=K$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=K$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=K$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L36', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$5=L$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$5=L$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$5=L$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$5=L$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$5=L$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$5=L$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$5=L$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$5=L$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$5=L$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$5=L$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$5=L$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$5=L$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$5=L$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$5=L$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$5=L$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$5=L$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$5=L$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L37', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$8=L$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$8=L$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$8=L$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$8=L$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$8=L$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$8=L$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$8=L$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$8=L$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$8=L$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$8=L$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$8=L$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$8=L$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$8=L$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$8=L$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$8=L$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$8=L$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$8=L$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L38', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$11=L$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$11=L$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$11=L$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$11=L$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$11=L$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$11=L$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$11=L$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$11=L$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$11=L$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$11=L$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$11=L$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$11=L$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$11=L$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$11=L$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$11=L$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$11=L$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$11=L$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L39', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$14=L$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$14=L$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$14=L$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$14=L$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$14=L$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$14=L$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$14=L$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$14=L$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$14=L$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$14=L$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$14=L$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$14=L$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$14=L$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$14=L$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$14=L$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$14=L$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$14=L$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L40', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$17=L$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$17=L$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$17=L$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$17=L$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$17=L$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$17=L$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$17=L$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$17=L$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$17=L$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$17=L$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$17=L$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$17=L$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$17=L$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$17=L$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$17=L$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$17=L$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$17=L$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"},
        {'cell': 'L41', 'formula': "=SUBSTITUTE(TRIM(CONCAT(IF(D111!$G$20=L$35,D111!$B$1 & \" \", \"\"),IF(D211!$G$20=L$35,D211!$B$1 & \" \", \"\"),IF(D311!$G$20=L$35,D311!$B$1 & \" \", \"\"),IF(D411!$G$20=L$35,D411!$B$1 & \" \", \"\"),IF(C111!$G$20=L$35,C111!$B$1 & \" \", \"\"),IF(C121!$G$20=L$35,C121!$B$1 & \" \", \"\"),IF(C122!$G$20=L$35,C122!$B$1 & \" \", \"\"),IF(C211!$G$20=L$35,C211!$B$1 & \" \", \"\"),IF(C212!$G$20=L$35,C212!$B$1 & \" \", \"\"),IF(C311!$G$20=L$35,C311!$B$1 & \" \", \"\"),IF(C312!$G$20=L$35,C312!$B$1 & \" \", \"\"),IF(C411!$G$20=L$35,C411!$B$1 & \" \", \"\"),IF(C412!$G$20=L$35,C412!$B$1 & \" \", \"\"),IF(M111!$G$20=L$35,M111!$B$1 & \" \", \"\"),IF(M211!$G$20=L$35,M211!$B$1 & \" \", \"\"),IF(M311!$G$20=L$35,M311!$B$1 & \" \", \"\"),IF(M411!$G$20=L$35,M411!$B$1 & \" \", \"\"))), \" \", \",\")"}
    ],
})


# Generar fórmulas dinámicas de Faltan y Asignadas para cada hoja de grupo
for sheet in sheets_cfg:
    if sheet.get('title') == 'Aulas':
        continue
    data = sheet.get('data', [])
    horario_range = sheet.get('horario_data_range', '$C$4:$G$15')
    formulas = sheet.get('formulas', [])
    for row_idx, row_data in enumerate(data):
        abrev = row_data[8] if len(row_data) > 8 else None
        if abrev and isinstance(abrev, str) and abrev.strip() and abrev.strip() != 'Abrev':
            excel_row = row_idx + 1
            formulas.append({'row': excel_row, 'col': 13, 'value': f'=COUNTIF({horario_range},I{excel_row})'})
            formulas.append({'row': excel_row, 'col': 12, 'value': f'=K{excel_row}-M{excel_row}'})
    sheet['formulas'] = formulas

config_excel = {'sheets': sheets_cfg}

generar_excel_personalizado(config_excel, 'propuesta_horarios_desde_lisp.xlsx')
