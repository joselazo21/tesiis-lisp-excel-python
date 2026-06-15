#!/usr/bin/env python3
from hoja_con_formulas import generar_excel_personalizado

config = {
    "sheets": [
        {
            "title": "Notas",
            "regions": [
{
        "data": [["", "", "", "", "", ""], ["", "", "", "", "", ""], ["Nombre", "Nota 1", "Nota 2", "Promedio", "", ""], ["Ana", "85", "90", "", "", ""], ["Luis", "70", "75", "", "", ""], ["Elena", "92", "88", "", "", ""], ["Carlos", "60", "65", "", "", ""], ["Sofía", "78", "82", "", "", ""], ["Pedro", "95", "91", "", "", ""]],
        "headers": ["", "", "", "", ""],
        "params": {"F1": 80},
        "formulas": [{"row": 4, "col": 4, "value": "=((B4+C4)/2)"}, {"row": 5, "col": 4, "value": "=((B5+C5)/2)"}, {"row": 6, "col": 4, "value": "=((B6+C6)/2)"}, {"row": 7, "col": 4, "value": "=((B7+C7)/2)"}, {"row": 8, "col": 4, "value": "=((B8+C8)/2)"}, {"row": 9, "col": 4, "value": "=((B9+C9)/2)"}],
        "conditional_formats": [{"range": "A4:A9", "formula": "D4>$F$1", "style": {"bg_color": "FF4444"}}, {"range": "B4:B9", "formula": "D4>$F$1", "style": {"bg_color": "FF4444"}}],
        "table_ranges": ["A4:F9"],
        "table_block_sizes": [{"range": "A4:D9", "row_step": 1, "col_step": 1}],
        "column_widths": {1: 10, 2: 10, 3: 10, 4: 10, 5: 0, 6: 8},
        "border_color": "4F81BD",
        "border_style": "thick"
    }            ]
        }    ]
}

generar_excel_personalizado(config, "Archivo-Excel.xlsx")

if __name__=='__main__':
    print('OK: Archivo-Excel.xlsx')
