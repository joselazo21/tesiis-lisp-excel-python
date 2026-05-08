#!/usr/bin/env python3
from hoja_con_formulas import generar_excel_personalizado

config = {#
    "sheets": [
        {
            "title": "Resumen",
        "data": [["TOTAL", 88, 3717]],
        "headers": ["", "Programas", "Minutos"],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "A2:C2", "style": {"bg_color": "FFF2CC"}}],
            "header_style": {"bold": True, "bg_color": "F4CCCC"}
        }
,
        {
            "title": " 1-Lunes",
        "data": [["16:00", "", 5, "EL TIEMPO Y LA MEMORIA", "informativo", "adulto"], ["16:05", "", 5, "COORDENADAS", "informativo", "adulto"], ["16:10", "", 50, "REVISTA HOLA HABANA", "revista", "adulto"], ["17:00", "", 30, "DÉCADAS MILAGROSAS", "musical", "adulto"], ["17:30", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["18:00", "", 30, "POWER RANGERS", "infantil", "infantil"], ["18:30", "", 25, "CINECITO EN TV", "cine", "infantil"], ["18:55", "", 5, "COORDENADAS INFANTILES", "informativo", "infantil"], ["19:00", "", 30, "VE Y MIRA", "cine", "adulto"], ["19:30", "", 30, "MÚSICA HABANA", "musical", "juvenil"], ["20:00", "", 30, "HABANA COLECCIÓN", "cultural", "adulto"], ["20:30", "", 60, "MÚSICA SÍ", "musical", "adulto"], ["21:30", "", 15, "D DISEÑO", "cultural", "adulto"], ["21:45", "", 15, "SIN PUNTOS SUSPENSIVOS", "entrevista", "adulto"], ["22:00", "", 45, "NOVELA “LA NIETA ELEGIDA”", "ficción", "adulto"], ["22:45", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["23:15", "", 45, "SERIE “CRÍMENES MAYORES”", "ficción", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}, {"row": 10, "col": 2, "value": "=TEXT(TIMEVALUE(A10)+(C10/1440),\"hh:mm\")"}, {"row": 11, "col": 2, "value": "=TEXT(TIMEVALUE(A11)+(C11/1440),\"hh:mm\")"}, {"row": 12, "col": 2, "value": "=TEXT(TIMEVALUE(A12)+(C12/1440),\"hh:mm\")"}, {"row": 13, "col": 2, "value": "=TEXT(TIMEVALUE(A13)+(C13/1440),\"hh:mm\")"}, {"row": 14, "col": 2, "value": "=TEXT(TIMEVALUE(A14)+(C14/1440),\"hh:mm\")"}, {"row": 15, "col": 2, "value": "=TEXT(TIMEVALUE(A15)+(C15/1440),\"hh:mm\")"}, {"row": 16, "col": 2, "value": "=TEXT(TIMEVALUE(A16)+(C16/1440),\"hh:mm\")"}, {"row": 17, "col": 2, "value": "=TEXT(TIMEVALUE(A17)+(C17/1440),\"hh:mm\")"}, {"row": 18, "col": 2, "value": "=TEXT(TIMEVALUE(A18)+(C18/1440),\"hh:mm\")"}, {"row": 19, "col": 2, "value": "=TEXT(TIMEVALUE(A19)+(C19/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "D9EAD3"}}, {"range": "E4:E4", "style": {"bg_color": "D9EAD3"}}, {"range": "E5:E5", "style": {"bg_color": "E6B8AF"}}, {"range": "E6:E6", "style": {"bg_color": "DDCDE4"}}, {"range": "E7:E7", "style": {"bg_color": "D9EAD3"}}, {"range": "E8:E8", "style": {"bg_color": "FFF2CC"}}, {"range": "E9:E9", "style": {"bg_color": "D9EAD3"}}, {"range": "E10:E10", "style": {"bg_color": "D9EAD3"}}, {"range": "E11:E11", "style": {"bg_color": "D9EAD3"}}, {"range": "E12:E12", "style": {"bg_color": "DDCDE4"}}, {"range": "E13:E13", "style": {"bg_color": "CFE2F3"}}, {"range": "E14:E14", "style": {"bg_color": "DDCDE4"}}, {"range": "E15:E15", "style": {"bg_color": "CFE2F3"}}, {"range": "E16:E16", "style": {"bg_color": "FCE5CD"}}, {"range": "E17:E17", "style": {"bg_color": "E6B8AF"}}, {"range": "E18:E18", "style": {"bg_color": "D9EAD3"}}, {"range": "E19:E19", "style": {"bg_color": "E6B8AF"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
,
        {
            "title": " 2-Martes",
        "data": [["19:30", "", 30, "TODO POP", "musical", "juvenil"], ["20:00", "", 30, "BREVES ESTACIONES", "cultural", "adulto"], ["20:30", "", 15, "SALUDARTE", "salud", "adulto"], ["20:45", "", 15, "SECUENCIA", "cultural", "adulto"], ["21:00", "", 30, "RITMO CLIP", "musical", "juvenil"], ["21:30", "", 30, "TRIANGULO DE LA CONFIANZA", "entrevista", "adulto"], ["22:00", "", 45, "NOVELA “LA NIETA ELEGIDA”", "ficción", "adulto"], ["22:45", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["23:15", "", 45, "SERIE “CRÍMENES MAYORES”", "ficción", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}, {"row": 10, "col": 2, "value": "=TEXT(TIMEVALUE(A10)+(C10/1440),\"hh:mm\")"}, {"row": 11, "col": 2, "value": "=TEXT(TIMEVALUE(A11)+(C11/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "DDCDE4"}}, {"range": "E4:E4", "style": {"bg_color": "CFE2F3"}}, {"range": "E5:E5", "style": {"bg_color": "C6EFCE"}}, {"range": "E6:E6", "style": {"bg_color": "CFE2F3"}}, {"range": "E7:E7", "style": {"bg_color": "DDCDE4"}}, {"range": "E8:E8", "style": {"bg_color": "FCE5CD"}}, {"range": "E9:E9", "style": {"bg_color": "E6B8AF"}}, {"range": "E10:E10", "style": {"bg_color": "D9EAD3"}}, {"range": "E11:E11", "style": {"bg_color": "E6B8AF"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
,
        {
            "title": " 3-Miércoles",
        "data": [["19:45", "", 15, "COSAS DEL CINE", "cine", "adulto"], ["20:00", "", 30, "RITMO CLIP", "musical", "juvenil"], ["20:30", "", 30, "BANDA SONORA JUVENIL", "cultural", "juvenil"], ["21:00", "", 30, "VE Y MIRA", "cine", "adulto"], ["21:30", "", 30, "MÚSICA DEL MUNDO", "musical", "adulto"], ["22:00", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["22:30", "", 77, "CINEMA HABANA", "cine", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "D9EAD3"}}, {"range": "E4:E4", "style": {"bg_color": "DDCDE4"}}, {"range": "E5:E5", "style": {"bg_color": "CFE2F3"}}, {"range": "E6:E6", "style": {"bg_color": "D9EAD3"}}, {"range": "E7:E7", "style": {"bg_color": "DDCDE4"}}, {"range": "E8:E8", "style": {"bg_color": "D9EAD3"}}, {"range": "E9:E9", "style": {"bg_color": "D9EAD3"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
,
        {
            "title": " 4-Jueves",
        "data": [["16:00", "", 5, "EL TIEMPO Y LA MEMORIA", "informativo", "adulto"], ["16:05", "", 5, "COORDENADAS", "informativo", "adulto"], ["16:10", "", 50, "REVISTA HOLA HABANA", "revista", "adulto"], ["17:00", "", 15, "SALUDARTE", "salud", "adulto"], ["17:15", "", 15, "SECUENCIA", "cultural", "adulto"], ["17:30", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["18:00", "", 30, "CAZADOR DE TROLES", "infantil", "infantil"], ["18:30", "", 15, "EL MUNDO DE CRAIG", "animacion", "infantil"], ["18:45", "", 45, "ÁNIMA", "animacion", "juvenil"], ["19:30", "", 30, "PAPEL EN BLANCO", "cultural", "adulto"], ["20:00", "", 30, "MÚSICA DEL MUNDO", "musical", "adulto"], ["20:30", "", 15, "TRAVESÍA", "cultural", "adulto"], ["20:45", "", 15, "DONDE VA LA HABANA", "cultural", "adulto"], ["21:00", "", 30, "BANDA SONORA", "musical", "adulto"], ["21:30", "", 30, "ESTA ES MI PEÑA", "musical", "adulto"], ["22:00", "", 45, "NOVELA “LA NIETA ELEGIDA”", "ficción", "adulto"], ["22:45", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["23:15", "", 45, "SERIE “CRÍMENES MAYORES”", "ficción", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}, {"row": 10, "col": 2, "value": "=TEXT(TIMEVALUE(A10)+(C10/1440),\"hh:mm\")"}, {"row": 11, "col": 2, "value": "=TEXT(TIMEVALUE(A11)+(C11/1440),\"hh:mm\")"}, {"row": 12, "col": 2, "value": "=TEXT(TIMEVALUE(A12)+(C12/1440),\"hh:mm\")"}, {"row": 13, "col": 2, "value": "=TEXT(TIMEVALUE(A13)+(C13/1440),\"hh:mm\")"}, {"row": 14, "col": 2, "value": "=TEXT(TIMEVALUE(A14)+(C14/1440),\"hh:mm\")"}, {"row": 15, "col": 2, "value": "=TEXT(TIMEVALUE(A15)+(C15/1440),\"hh:mm\")"}, {"row": 16, "col": 2, "value": "=TEXT(TIMEVALUE(A16)+(C16/1440),\"hh:mm\")"}, {"row": 17, "col": 2, "value": "=TEXT(TIMEVALUE(A17)+(C17/1440),\"hh:mm\")"}, {"row": 18, "col": 2, "value": "=TEXT(TIMEVALUE(A18)+(C18/1440),\"hh:mm\")"}, {"row": 19, "col": 2, "value": "=TEXT(TIMEVALUE(A19)+(C19/1440),\"hh:mm\")"}, {"row": 20, "col": 2, "value": "=TEXT(TIMEVALUE(A20)+(C20/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "D9EAD3"}}, {"range": "E4:E4", "style": {"bg_color": "D9EAD3"}}, {"range": "E5:E5", "style": {"bg_color": "E6B8AF"}}, {"range": "E6:E6", "style": {"bg_color": "C6EFCE"}}, {"range": "E7:E7", "style": {"bg_color": "CFE2F3"}}, {"range": "E8:E8", "style": {"bg_color": "D9EAD3"}}, {"range": "E9:E9", "style": {"bg_color": "FFF2CC"}}, {"range": "E10:E10", "style": {"bg_color": "E6B8AF"}}, {"range": "E11:E11", "style": {"bg_color": "E6B8AF"}}, {"range": "E12:E12", "style": {"bg_color": "CFE2F3"}}, {"range": "E13:E13", "style": {"bg_color": "DDCDE4"}}, {"range": "E14:E14", "style": {"bg_color": "CFE2F3"}}, {"range": "E15:E15", "style": {"bg_color": "CFE2F3"}}, {"range": "E16:E16", "style": {"bg_color": "DDCDE4"}}, {"range": "E17:E17", "style": {"bg_color": "DDCDE4"}}, {"range": "E18:E18", "style": {"bg_color": "E6B8AF"}}, {"range": "E19:E19", "style": {"bg_color": "D9EAD3"}}, {"range": "E20:E20", "style": {"bg_color": "E6B8AF"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
,
        {
            "title": " 5-Viernes",
        "data": [["16:30", "", 60, "GEN HABANERO", "documental", "adulto"], ["17:30", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["18:00", "", 115, "TIENE QUE VER", "cine", "infantil"], ["19:55", "", 5, "COORDENADAS INFANTILES", "informativo", "infantil"], ["20:00", "", 30, "ALGO ENTRE MANOS", "cultural", "adulto"], ["20:30", "", 30, "BREVES ESTACIONES", "cultural", "adulto"], ["21:00", "", 30, "DÉCADAS MILAGROSAS", "musical", "adulto"], ["21:30", "", 15, "SIN PUNTOS SUSPENSIVOS", "entrevista", "adulto"], ["21:45", "", 15, "YO BAILO", "musical", "toda-la-familia"], ["22:00", "", 45, "NOVELA “LA NIETA ELEGIDA”", "ficción", "adulto"], ["22:45", "", 30, "HABANA NOTICIARIO", "informativo", "adulto"], ["23:15", "", 45, "SERIE “CRÍMENES MAYORES”", "ficción", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}, {"row": 10, "col": 2, "value": "=TEXT(TIMEVALUE(A10)+(C10/1440),\"hh:mm\")"}, {"row": 11, "col": 2, "value": "=TEXT(TIMEVALUE(A11)+(C11/1440),\"hh:mm\")"}, {"row": 12, "col": 2, "value": "=TEXT(TIMEVALUE(A12)+(C12/1440),\"hh:mm\")"}, {"row": 13, "col": 2, "value": "=TEXT(TIMEVALUE(A13)+(C13/1440),\"hh:mm\")"}, {"row": 14, "col": 2, "value": "=TEXT(TIMEVALUE(A14)+(C14/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "FFF2CC"}}, {"range": "E4:E4", "style": {"bg_color": "D9EAD3"}}, {"range": "E5:E5", "style": {"bg_color": "D9EAD3"}}, {"range": "E6:E6", "style": {"bg_color": "D9EAD3"}}, {"range": "E7:E7", "style": {"bg_color": "CFE2F3"}}, {"range": "E8:E8", "style": {"bg_color": "CFE2F3"}}, {"range": "E9:E9", "style": {"bg_color": "DDCDE4"}}, {"range": "E10:E10", "style": {"bg_color": "FCE5CD"}}, {"range": "E11:E11", "style": {"bg_color": "DDCDE4"}}, {"range": "E12:E12", "style": {"bg_color": "E6B8AF"}}, {"range": "E13:E13", "style": {"bg_color": "D9EAD3"}}, {"range": "E14:E14", "style": {"bg_color": "E6B8AF"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
,
        {
            "title": " 6-Sábado",
        "data": [["16:00", "", 5, "EL TIEMPO Y LA MEMORIA", "informativo", "adulto"], ["16:05", "", 5, "COORDENADAS", "informativo", "adulto"], ["16:10", "", 35, "CUENTAS VERDE LIMÓN", "infantil", "infantil"], ["16:45", "", 30, "BANDA SONORA JUVENIL", "cine", "juvenil"], ["17:15", "", 15, "TRAVESÍA", "cultural", "adulto"], ["17:30", "", 30, "TODO POP", "musical", "juvenil"], ["18:00", "", 45, "SERIE JUVENIL “SABRINA, LA BRUJA ADOLESCENTE”", "ficción", "juvenil"], ["18:45", "", 45, "LIBRE ACCESO", "informativo", "adulto"], ["19:30", "", 30, "JUGADA PERFECTA", "deporte", "toda-la-familia"], ["20:00", "", 15, "QUE LA MÚSICA NO FALTE", "musical", "adulto"], ["20:15", "", 15, "COSAS DEL CINE", "cine", "adulto"], ["20:30", "", 30, "MÚSICA HABANA", "musical", "adulto"], ["21:00", "", 90, "X DISTANTE", "animacion", "juvenil"], ["22:30", "", 90, "MÚSICA SI", "musical", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}, {"row": 10, "col": 2, "value": "=TEXT(TIMEVALUE(A10)+(C10/1440),\"hh:mm\")"}, {"row": 11, "col": 2, "value": "=TEXT(TIMEVALUE(A11)+(C11/1440),\"hh:mm\")"}, {"row": 12, "col": 2, "value": "=TEXT(TIMEVALUE(A12)+(C12/1440),\"hh:mm\")"}, {"row": 13, "col": 2, "value": "=TEXT(TIMEVALUE(A13)+(C13/1440),\"hh:mm\")"}, {"row": 14, "col": 2, "value": "=TEXT(TIMEVALUE(A14)+(C14/1440),\"hh:mm\")"}, {"row": 15, "col": 2, "value": "=TEXT(TIMEVALUE(A15)+(C15/1440),\"hh:mm\")"}, {"row": 16, "col": 2, "value": "=TEXT(TIMEVALUE(A16)+(C16/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "D9EAD3"}}, {"range": "E4:E4", "style": {"bg_color": "D9EAD3"}}, {"range": "E5:E5", "style": {"bg_color": "FFF2CC"}}, {"range": "E6:E6", "style": {"bg_color": "D9EAD3"}}, {"range": "E7:E7", "style": {"bg_color": "CFE2F3"}}, {"range": "E8:E8", "style": {"bg_color": "DDCDE4"}}, {"range": "E9:E9", "style": {"bg_color": "E6B8AF"}}, {"range": "E10:E10", "style": {"bg_color": "D9EAD3"}}, {"range": "E11:E11", "style": {"bg_color": "C6EFCE"}}, {"range": "E12:E12", "style": {"bg_color": "DDCDE4"}}, {"range": "E13:E13", "style": {"bg_color": "D9EAD3"}}, {"range": "E14:E14", "style": {"bg_color": "DDCDE4"}}, {"range": "E15:E15", "style": {"bg_color": "E6B8AF"}}, {"range": "E16:E16", "style": {"bg_color": "DDCDE4"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
,
        {
            "title": " 7-Domingo",
        "data": [["02:00", "", 5, "EL TIEMPO Y LA MEMORIA", "informativo", "adulto"], ["02:05", "", 5, "COORDENADAS", "informativo", "adulto"], ["02:10", "", 890, "CANAL HABANA DEPORTES", "deporte", "toda-la-familia"], ["17:00", "", 60, "LATINOS", "musical", "adulto"], ["18:00", "", 30, "PAPEL EN BLANCO", "cultural", "adulto"], ["18:30", "", 30, "ALGO ENTRE MANOS", "cultural", "adulto"], ["19:00", "", 15, "VERDE HABANA", "documental", "toda-la-familia"], ["19:15", "", 15, "GEN HABANERO", "documental", "adulto"], ["19:30", "", 30, "TRIANGULO DE LA CONFIANZA", "entrevista", "adulto"], ["20:00", "", 30, "BANDA SONORA", "musical", "adulto"], ["20:30", "", 205, "CINE +", "cine", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}, {"row": 4, "col": 2, "value": "=TEXT(TIMEVALUE(A4)+(C4/1440),\"hh:mm\")"}, {"row": 5, "col": 2, "value": "=TEXT(TIMEVALUE(A5)+(C5/1440),\"hh:mm\")"}, {"row": 6, "col": 2, "value": "=TEXT(TIMEVALUE(A6)+(C6/1440),\"hh:mm\")"}, {"row": 7, "col": 2, "value": "=TEXT(TIMEVALUE(A7)+(C7/1440),\"hh:mm\")"}, {"row": 8, "col": 2, "value": "=TEXT(TIMEVALUE(A8)+(C8/1440),\"hh:mm\")"}, {"row": 9, "col": 2, "value": "=TEXT(TIMEVALUE(A9)+(C9/1440),\"hh:mm\")"}, {"row": 10, "col": 2, "value": "=TEXT(TIMEVALUE(A10)+(C10/1440),\"hh:mm\")"}, {"row": 11, "col": 2, "value": "=TEXT(TIMEVALUE(A11)+(C11/1440),\"hh:mm\")"}, {"row": 12, "col": 2, "value": "=TEXT(TIMEVALUE(A12)+(C12/1440),\"hh:mm\")"}, {"row": 13, "col": 2, "value": "=TEXT(TIMEVALUE(A13)+(C13/1440),\"hh:mm\")"}],
            "table_borders": True,
            "border_color": "B7B7B7",
            "range_styles": [{"range": "E3:E3", "style": {"bg_color": "D9EAD3"}}, {"range": "E4:E4", "style": {"bg_color": "D9EAD3"}}, {"range": "E5:E5", "style": {"bg_color": "C6EFCE"}}, {"range": "E6:E6", "style": {"bg_color": "DDCDE4"}}, {"range": "E7:E7", "style": {"bg_color": "CFE2F3"}}, {"range": "E8:E8", "style": {"bg_color": "CFE2F3"}}, {"range": "E9:E9", "style": {"bg_color": "FFF2CC"}}, {"range": "E10:E10", "style": {"bg_color": "FFF2CC"}}, {"range": "E11:E11", "style": {"bg_color": "FCE5CD"}}, {"range": "E12:E12", "style": {"bg_color": "DDCDE4"}}, {"range": "E13:E13", "style": {"bg_color": "D9EAD3"}}],
            "column_widths": {10: 8, 8: 30, 12: 10},
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
    ]
}

generar_excel_personalizado(config, "Canal Habana.xlsx")

if __name__=='__main__':
    print('OK: Canal Habana.xlsx')
