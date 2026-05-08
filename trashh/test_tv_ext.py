#!/usr/bin/env python3
from hoja_con_formulas import generar_excel_personalizado

config = {#
    "sheets": [
        {
            "title": " 1-Lunes",
        "data": [["16:00", "16:30", 30, "Test Program", "informativo", "adulto"]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "table_borders": True,
            "border_color": "B7B7B7",
            "border_style": "thin",
            "header_style": {"bold": True, "bg_color": "4A90E2", "align": "center"}
        }
    ]
}

generar_excel_personalizado(config, "Test.xlsx")

if __name__=='__main__':
    print('OK: Test.xlsx')
