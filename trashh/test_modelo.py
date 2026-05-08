#!/usr/bin/env python3
from hoja_con_formulas import generar_excel_personalizado

config = {#
    "sheets": [
        {
            "title": "Test",
        "data": [["Program A", 30], ["Program B", 45]],
        "headers": ["Programa", "Duracion"],
            "table_borders": True,
            "border_color": "B7B7B7",
            "border_style": "thin",
            "header_style": {"bold": True, "bg_color": "D9EAD3"}
        }
    ]
}

generar_excel_personalizado(config, "Test.xlsx")

if __name__=='__main__':
    print('OK: Test.xlsx')
