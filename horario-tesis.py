#!/usr/bin/env python3
from hoja_con_formulas import generar_excel_personalizado

config = {
    "sheets": [
        {
            "title": "Defensas",
            "regions": [
{
        "data": [["", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["Estudiante", "Tutor", "Oponente", "Presidente", "Vocal", "Secretario", "Dia", "Hora", "Local", "", "Nombre", "Grado", "Defensas", ""], ["Juan Diaz", "Piad", "Garcia", "Lopez", "Torres", "Ruiz", "Lunes", "10:00", "Aula 1", "", "Piad", "Dr", "", ""], ["Maria Vega", "Garcia", "Lopez", "Piad", "Soto", "Torres", "Lunes", "14:00", "Aula 2", "", "Garcia", "Dr", "", ""], ["Luis Mora", "Lopez", "Torres", "Garcia", "Ruiz", "Soto", "Lunes", "10:00", "Aula 3", "", "Lopez", "Msc", "", ""], ["", "", "", "", "", "", "", "", "", "", "Torres", "Msc", "", ""], ["", "", "", "", "", "", "", "", "", "", "Ruiz", "Lic", "", ""], ["", "", "", "", "", "", "", "", "", "", "Soto", "Msc", "", ""]],
        "headers": ["", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        "formulas": [{"row": 4, "col": 13, "value": "=COUNTIF($B$4:$F$6,K4)"}, {"row": 5, "col": 13, "value": "=COUNTIF($B$4:$F$6,K5)"}, {"row": 6, "col": 13, "value": "=COUNTIF($B$4:$F$6,K6)"}, {"row": 7, "col": 1, "value": "=\"Total defensas:\""}, {"row": 7, "col": 2, "value": "=COUNTA($A$4:$A$6)"}, {"row": 7, "col": 13, "value": "=COUNTIF($B$4:$F$6,K7)"}, {"row": 8, "col": 13, "value": "=COUNTIF($B$4:$F$6,K8)"}, {"row": 9, "col": 13, "value": "=COUNTIF($B$4:$F$6,K9)"}, {"row": 10, "col": 11, "value": "=\"Total profesores:\""}, {"row": 10, "col": 13, "value": "=SUM($M$4:$M$9)"}],
        "conditional_formats": [{"range": "K4:K9", "formula": "SUMPRODUCT((($B$4:$B$6=$K4)+($C$4:$C$6=$K4)+($D$4:$D$6=$K4)+($E$4:$E$6=$K4)+($F$4:$F$6=$K4))*(COUNTIFS($G$4:$G$6,$G$4:$G$6,$H$4:$H$6,$H$4:$H$6)>1))>0", "style": {"font_color": "#FF0000"}}],
        "table_ranges": ["A4:N9"],
        "table_block_sizes": [{"range": "A4:I6", "row_step": 1, "col_step": 1}, {"range": "K4:M9", "row_step": 1, "col_step": 1}],
        "column_widths": {1: 10, 2: 10, 3: 10, 4: 10, 5: 10, 6: 10, 7: 10, 8: 10, 9: 10, 10: 0, 11: 10, 12: 10, 13: 10, 14: 0},
        "border_color": "4F81BD",
        "border_style": "thick"
    }            ]
        }    ]
}

generar_excel_personalizado(config, "Horario_Tesis.xlsx")

if __name__=='__main__':
    print('OK: Horario_Tesis.xlsx')
