#!/usr/bin/env python3
from hoja_con_formulas import generar_excel_personalizado

config = {
    "sheets": [
        {
            "title": "Resumen",
            "regions": [
{        "data": [["TOTAL", 1, 30]],
        "headers": ["", "Programas", "Minutos"],
            "range_styles": [{"range": "A2:C2", "style": {"bg_color": "FFF2CC"}}],
}            ]
        },
        {
            "title": " 1-Lunes",
            "regions": [
{        "data": [["", "", 30, "Prog A", "", ""]],
        "headers": ["Inicio", "Fin", "Duración", "Programa", "Tipo", "Público"],
            "formulas": [{"row": 3, "col": 2, "value": "=TEXT(TIMEVALUE(A3)+(C3/1440),\"hh:mm\")"}],
            "column_widths": {10: 8, 8: 30, 12: 10},
}            ]
        }    ]
}

generar_excel_personalizado(config, "Canal Habana.xlsx")

if __name__=='__main__':
    print('OK: Canal Habana.xlsx')
