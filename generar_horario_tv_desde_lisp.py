from hoja_con_formulas import generar_excel_horario_tv_desde_parametros

nombre_canal = "Canal Habana"
planificacion_semanal = [
    {'dia': "lunes", 'programas': [
        {'nombre': "EL TIEMPO Y LA MEMORIA", 'duracion': 5, 'hora_inicio': "16:00", 'hora_final': "16:05", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "COORDENADAS", 'duracion': 5, 'hora_inicio': "16:05", 'hora_final': "16:10", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "REVISTA HOLA HABANA", 'duracion': 50, 'hora_inicio': "16:10", 'hora_final': "17:00", 'tipo_programa': "revista", 'tipo_publico': "adulto"},
        {'nombre': "DÉCADAS MILAGROSAS", 'duracion': 30, 'hora_inicio': "17:00", 'hora_final': "17:30", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "17:30", 'hora_final': "18:00", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "POWER RANGERS", 'duracion': 30, 'hora_inicio': "18:00", 'hora_final': "18:30", 'tipo_programa': "infantil", 'tipo_publico': "infantil"},
        {'nombre': "CINECITO EN TV", 'duracion': 25, 'hora_inicio': "18:30", 'hora_final': "18:55", 'tipo_programa': "cine", 'tipo_publico': "infantil"},
        {'nombre': "COORDENADAS INFANTILES", 'duracion': 5, 'hora_inicio': "18:55", 'hora_final': "19:00", 'tipo_programa': "informativo", 'tipo_publico': "infantil"},
        {'nombre': "VE Y MIRA", 'duracion': 30, 'hora_inicio': "19:00", 'hora_final': "19:30", 'tipo_programa': "cine", 'tipo_publico': "adulto"},
        {'nombre': "MÚSICA HABANA", 'duracion': 30, 'hora_inicio': "19:30", 'hora_final': "20:00", 'tipo_programa': "musical", 'tipo_publico': "juvenil"},
        {'nombre': "HABANA COLECCIÓN", 'duracion': 30, 'hora_inicio': "20:00", 'hora_final': "20:30", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "MÚSICA SÍ", 'duracion': 60, 'hora_inicio': "20:30", 'hora_final': "21:30", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "D DISEÑO", 'duracion': 15, 'hora_inicio': "21:30", 'hora_final': "21:45", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "SIN PUNTOS SUSPENSIVOS", 'duracion': 15, 'hora_inicio': "21:45", 'hora_final': "22:00", 'tipo_programa': "entrevista", 'tipo_publico': "adulto"},
        {'nombre': "NOVELA “LA NIETA ELEGIDA”", 'duracion': 45, 'hora_inicio': "22:00", 'hora_final': "22:45", 'tipo_programa': "ficción", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "22:45", 'hora_final': "23:15", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "SERIE “CRÍMENES MAYORES”", 'duracion': 45, 'hora_inicio': "23:15", 'hora_final': "00:00", 'tipo_programa': "ficción", 'tipo_publico': "adulto"}
    ]},
    {'dia': "martes", 'programas': [
        {'nombre': "TODO POP", 'duracion': 30, 'hora_inicio': "19:30", 'hora_final': "20:00", 'tipo_programa': "musical", 'tipo_publico': "juvenil"},
        {'nombre': "BREVES ESTACIONES", 'duracion': 30, 'hora_inicio': "20:00", 'hora_final': "20:30", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "SALUDARTE", 'duracion': 15, 'hora_inicio': "20:30", 'hora_final': "20:45", 'tipo_programa': "salud", 'tipo_publico': "adulto"},
        {'nombre': "SECUENCIA", 'duracion': 15, 'hora_inicio': "20:45", 'hora_final': "21:00", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "RITMO CLIP", 'duracion': 30, 'hora_inicio': "21:00", 'hora_final': "21:30", 'tipo_programa': "musical", 'tipo_publico': "juvenil"},
        {'nombre': "TRIANGULO DE LA CONFIANZA", 'duracion': 30, 'hora_inicio': "21:30", 'hora_final': "22:00", 'tipo_programa': "entrevista", 'tipo_publico': "adulto"},
        {'nombre': "NOVELA “LA NIETA ELEGIDA”", 'duracion': 45, 'hora_inicio': "22:00", 'hora_final': "22:45", 'tipo_programa': "ficción", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "22:45", 'hora_final': "23:15", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "SERIE “CRÍMENES MAYORES”", 'duracion': 45, 'hora_inicio': "23:15", 'hora_final': "00:00", 'tipo_programa': "ficción", 'tipo_publico': "adulto"}
    ]},
    {'dia': "miercoles", 'programas': [
        {'nombre': "COSAS DEL CINE", 'duracion': 15, 'hora_inicio': "19:45", 'hora_final': "20:00", 'tipo_programa': "cine", 'tipo_publico': "adulto"},
        {'nombre': "RITMO CLIP", 'duracion': 30, 'hora_inicio': "20:00", 'hora_final': "20:30", 'tipo_programa': "musical", 'tipo_publico': "juvenil"},
        {'nombre': "BANDA SONORA JUVENIL", 'duracion': 30, 'hora_inicio': "20:30", 'hora_final': "21:00", 'tipo_programa': "cultural", 'tipo_publico': "juvenil"},
        {'nombre': "VE Y MIRA", 'duracion': 30, 'hora_inicio': "21:00", 'hora_final': "21:30", 'tipo_programa': "cine", 'tipo_publico': "adulto"},
        {'nombre': "MÚSICA DEL MUNDO", 'duracion': 30, 'hora_inicio': "21:30", 'hora_final': "22:00", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "22:00", 'hora_final': "22:30", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "CINEMA HABANA", 'duracion': 77, 'hora_inicio': "22:30", 'hora_final': "23:47", 'tipo_programa': "cine", 'tipo_publico': "adulto"}
    ]},
    {'dia': "jueves", 'programas': [
        {'nombre': "EL TIEMPO Y LA MEMORIA", 'duracion': 5, 'hora_inicio': "16:00", 'hora_final': "16:05", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "COORDENADAS", 'duracion': 5, 'hora_inicio': "16:05", 'hora_final': "16:10", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "REVISTA HOLA HABANA", 'duracion': 50, 'hora_inicio': "16:10", 'hora_final': "17:00", 'tipo_programa': "revista", 'tipo_publico': "adulto"},
        {'nombre': "SALUDARTE", 'duracion': 15, 'hora_inicio': "17:00", 'hora_final': "17:15", 'tipo_programa': "salud", 'tipo_publico': "adulto"},
        {'nombre': "SECUENCIA", 'duracion': 15, 'hora_inicio': "17:15", 'hora_final': "17:30", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "17:30", 'hora_final': "18:00", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "CAZADOR DE TROLES", 'duracion': 30, 'hora_inicio': "18:00", 'hora_final': "18:30", 'tipo_programa': "infantil", 'tipo_publico': "infantil"},
        {'nombre': "EL MUNDO DE CRAIG", 'duracion': 15, 'hora_inicio': "18:30", 'hora_final': "18:45", 'tipo_programa': "animacion", 'tipo_publico': "infantil"},
        {'nombre': "ÁNIMA", 'duracion': 45, 'hora_inicio': "18:45", 'hora_final': "19:30", 'tipo_programa': "animacion", 'tipo_publico': "juvenil"},
        {'nombre': "PAPEL EN BLANCO", 'duracion': 30, 'hora_inicio': "19:30", 'hora_final': "20:00", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "MÚSICA DEL MUNDO", 'duracion': 30, 'hora_inicio': "20:00", 'hora_final': "20:30", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "TRAVESÍA", 'duracion': 15, 'hora_inicio': "20:30", 'hora_final': "20:45", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "DONDE VA LA HABANA", 'duracion': 15, 'hora_inicio': "20:45", 'hora_final': "21:00", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "BANDA SONORA", 'duracion': 30, 'hora_inicio': "21:00", 'hora_final': "21:30", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "ESTA ES MI PEÑA", 'duracion': 30, 'hora_inicio': "21:30", 'hora_final': "22:00", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "NOVELA “LA NIETA ELEGIDA”", 'duracion': 45, 'hora_inicio': "22:00", 'hora_final': "22:45", 'tipo_programa': "ficción", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "22:45", 'hora_final': "23:15", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "SERIE “CRÍMENES MAYORES”", 'duracion': 45, 'hora_inicio': "23:15", 'hora_final': "00:00", 'tipo_programa': "ficción", 'tipo_publico': "adulto"}
    ]},
    {'dia': "viernes", 'programas': [
        {'nombre': "GEN HABANERO", 'duracion': 60, 'hora_inicio': "16:30", 'hora_final': "17:30", 'tipo_programa': "documental", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "17:30", 'hora_final': "18:00", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "TIENE QUE VER", 'duracion': 115, 'hora_inicio': "18:00", 'hora_final': "19:55", 'tipo_programa': "cine", 'tipo_publico': "infantil"},
        {'nombre': "COORDENADAS INFANTILES", 'duracion': 5, 'hora_inicio': "19:55", 'hora_final': "20:00", 'tipo_programa': "informativo", 'tipo_publico': "infantil"},
        {'nombre': "ALGO ENTRE MANOS", 'duracion': 30, 'hora_inicio': "20:00", 'hora_final': "20:30", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "BREVES ESTACIONES", 'duracion': 30, 'hora_inicio': "20:30", 'hora_final': "21:00", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "DÉCADAS MILAGROSAS", 'duracion': 30, 'hora_inicio': "21:00", 'hora_final': "21:30", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "SIN PUNTOS SUSPENSIVOS", 'duracion': 15, 'hora_inicio': "21:30", 'hora_final': "21:45", 'tipo_programa': "entrevista", 'tipo_publico': "adulto"},
        {'nombre': "YO BAILO", 'duracion': 15, 'hora_inicio': "21:45", 'hora_final': "22:00", 'tipo_programa': "musical", 'tipo_publico': "toda-la-familia"},
        {'nombre': "NOVELA “LA NIETA ELEGIDA”", 'duracion': 45, 'hora_inicio': "22:00", 'hora_final': "22:45", 'tipo_programa': "ficción", 'tipo_publico': "adulto"},
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "22:45", 'hora_final': "23:15", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "SERIE “CRÍMENES MAYORES”", 'duracion': 45, 'hora_inicio': "23:15", 'hora_final': "00:00", 'tipo_programa': "ficción", 'tipo_publico': "adulto"}
    ]},
    {'dia': "sabado", 'programas': [
        {'nombre': "EL TIEMPO Y LA MEMORIA", 'duracion': 5, 'hora_inicio': "16:00", 'hora_final': "16:05", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "COORDENADAS", 'duracion': 5, 'hora_inicio': "16:05", 'hora_final': "16:10", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "CUENTAS VERDE LIMÓN", 'duracion': 35, 'hora_inicio': "16:10", 'hora_final': "16:45", 'tipo_programa': "infantil", 'tipo_publico': "infantil"},
        {'nombre': "BANDA SONORA JUVENIL", 'duracion': 30, 'hora_inicio': "16:45", 'hora_final': "17:15", 'tipo_programa': "cine", 'tipo_publico': "juvenil"},
        {'nombre': "TRAVESÍA", 'duracion': 15, 'hora_inicio': "17:15", 'hora_final': "17:30", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "TODO POP", 'duracion': 30, 'hora_inicio': "17:30", 'hora_final': "18:00", 'tipo_programa': "musical", 'tipo_publico': "juvenil"},
        {'nombre': "SERIE JUVENIL “SABRINA, LA BRUJA ADOLESCENTE”", 'duracion': 45, 'hora_inicio': "18:00", 'hora_final': "18:45", 'tipo_programa': "ficción", 'tipo_publico': "juvenil"},
        {'nombre': "LIBRE ACCESO", 'duracion': 45, 'hora_inicio': "18:45", 'hora_final': "19:30", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "JUGADA PERFECTA", 'duracion': 30, 'hora_inicio': "19:30", 'hora_final': "20:00", 'tipo_programa': "deporte", 'tipo_publico': "toda-la-familia"},
        {'nombre': "QUE LA MÚSICA NO FALTE", 'duracion': 15, 'hora_inicio': "20:00", 'hora_final': "20:15", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "COSAS DEL CINE", 'duracion': 15, 'hora_inicio': "20:15", 'hora_final': "20:30", 'tipo_programa': "cine", 'tipo_publico': "adulto"},
        {'nombre': "MÚSICA HABANA", 'duracion': 30, 'hora_inicio': "20:30", 'hora_final': "21:00", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "X DISTANTE", 'duracion': 90, 'hora_inicio': "21:00", 'hora_final': "22:30", 'tipo_programa': "animacion", 'tipo_publico': "juvenil"},
        {'nombre': "MÚSICA SI", 'duracion': 90, 'hora_inicio': "22:30", 'hora_final': "00:00", 'tipo_programa': "musical", 'tipo_publico': "adulto"}
    ]},
    {'dia': "domingo", 'programas': [
        {'nombre': "EL TIEMPO Y LA MEMORIA", 'duracion': 5, 'hora_inicio': "02:00", 'hora_final': "02:05", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "COORDENADAS", 'duracion': 5, 'hora_inicio': "02:05", 'hora_final': "02:10", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "CANAL HABANA DEPORTES", 'duracion': 890, 'hora_inicio': "02:10", 'hora_final': "17:00", 'tipo_programa': "deporte", 'tipo_publico': "toda-la-familia"},
        {'nombre': "LATINOS", 'duracion': 60, 'hora_inicio': "17:00", 'hora_final': "18:00", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "PAPEL EN BLANCO", 'duracion': 30, 'hora_inicio': "18:00", 'hora_final': "18:30", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "ALGO ENTRE MANOS", 'duracion': 30, 'hora_inicio': "18:30", 'hora_final': "19:00", 'tipo_programa': "cultural", 'tipo_publico': "adulto"},
        {'nombre': "VERDE HABANA", 'duracion': 15, 'hora_inicio': "19:00", 'hora_final': "19:15", 'tipo_programa': "documental", 'tipo_publico': "toda-la-familia"},
        {'nombre': "GEN HABANERO", 'duracion': 15, 'hora_inicio': "19:15", 'hora_final': "19:30", 'tipo_programa': "documental", 'tipo_publico': "adulto"},
        {'nombre': "TRIANGULO DE LA CONFIANZA", 'duracion': 30, 'hora_inicio': "19:30", 'hora_final': "20:00", 'tipo_programa': "entrevista", 'tipo_publico': "adulto"},
        {'nombre': "BANDA SONORA", 'duracion': 30, 'hora_inicio': "20:00", 'hora_final': "20:30", 'tipo_programa': "musical", 'tipo_publico': "adulto"},
        {'nombre': "CINE +", 'duracion': 205, 'hora_inicio': "20:30", 'hora_final': "23:55", 'tipo_programa': "cine", 'tipo_publico': "adulto"}
    ]}
]

generar_excel_horario_tv_desde_parametros(
    filename="horario_tv_semanal.xlsx",
    nombre_canal=nombre_canal,
    planificacion_semanal=planificacion_semanal,
    incluir_resumen=True
)
