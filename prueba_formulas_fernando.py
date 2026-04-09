#!/usr/bin/env python3
"""
Script de prueba: Inyecta fórmulas de Fernando en un Excel
Demuestra cómo usar las fórmulas complejas del ODS en el flujo Python/LISP
"""

import json
from hoja_con_formulas import generate_excel

def cargar_formulas_fernando(locale='es'):
    """Carga las fórmulas convertidas desde el JSON
    
    Args:
        locale: 'es' para español (default), 'en' para inglés
    """
    try:
        filename = f'formulas_fernando_convertidas_{locale}.json'
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ Archivo formulas_fernando_convertidas_{locale}.json no encontrado")
        return {}

def generar_excel_con_formulas_fernando(locale='es'):
    """Genera un Excel de prueba con las fórmulas complejas de Fernando
    
    Args:
        locale: 'es' para español, 'en' para inglés
    """
    
    print("=" * 80)
    print(f"GENERANDO EXCEL CON FÓRMULAS DE FERNANDO ({locale.upper()})")
    print("=" * 80)
    
    # Cargar fórmulas
    formulas_data = cargar_formulas_fernando(locale=locale)
    
    if not formulas_data:
        print("⚠️  No hay fórmulas cargadas, usando valores de demostración")
        formulas_data = {}
    
    # Preparar fórmulas para la hoja Aulas (las más complejas)
    aulas_formulas = []
    if 'Aulas' in formulas_data:
        print(f"\n✅ Cargadas {len(formulas_data['Aulas'])} fórmulas de la hoja Aulas")
        # Convertir formato JSON a formato esperado por la función
        for formula_item in formulas_data['Aulas'][:10]:  # Primeras 10 como ejemplo
            aulas_formulas.append({
                'cell': formula_item.get('cell'),
                'formula': formula_item.get('excel')
            })
    
    # Configuración Excel
    config = {
        "sheets": [
            # Hoja C111 con fórmulas simples
            {
                "title": "C111",
                "headers": ["Asignatura", "Profesor", "Aula L", "Aula M", "Aula V", 
                           "Estudiantes", "Faltan", "Frecuencia", "Asignadas"],
                "data": [
                    ["Análisis I", "Dr. García", "C111-1", "C111-2", "C111-3", 30, 5, 10, 25],
                    ["Álgebra", "Dra. López", "C111-4", "C111-5", "C111-6", 28, 3, 8, 25],
                    ["Cálculo", "Dr. Martínez", "C111-7", "C111-8", "C111-9", 32, 7, 12, 25],
                ],
                "formulas": [
                    {"row": 2, "col": 7, "value": "=D2-E2"},  # Faltan = Estudiantes - Asignadas
                    {"row": 3, "col": 7, "value": "=D3-E3"},
                    {"row": 4, "col": 7, "value": "=D4-E4"},
                ],
                "column_widths": {1: 15, 2: 15, 3: 12, 4: 12, 5: 12, 6: 12, 7: 10, 8: 12, 9: 12},
                "header_style": {"bold": True, "align": "center", "bg_color": "4472C4", "color": "FFFFFF"},
                "table_borders": True,
                "border_color": "000000",
                "border_style": "thin",
                "table_ranges": ["A1:I4"],
            },
            # Hoja Aulas con fórmulas COMPLEJAS
            {
                "title": "Aulas",
                "headers": ["Aula", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
                "data": [
                    ["A101", "", "", "", "", ""],
                    ["A102", "", "", "", "", ""],
                    ["A103", "", "", "", "", ""],
                    ["A104", "", "", "", "", ""],
                    ["A105", "", "", "", "", ""],
                ],
                "fernando_formulas": aulas_formulas,  # ✅ INYECTAR FÓRMULAS DE FERNANDO
                "column_widths": {1: 12, 2: 25, 3: 25, 4: 25, 5: 25, 6: 25},
                "header_style": {"bold": True, "align": "center", "bg_color": "70AD47", "color": "FFFFFF"},
                "table_borders": True,
                "border_color": "000000",
                "border_style": "thin",
                "table_ranges": ["A1:F6"],
            }
        ]
    }
    
    # Generar Excel
    output_file = f"prueba_formulas_fernando_{locale}.xlsx"
    generate_excel(config, output_file)
    
    print(f"\n✅ Archivo generado: {output_file}")
    print("\nContenido generado:")
    print(f"  • Hoja 'C111': Datos con fórmulas simples")
    print(f"  • Hoja 'Aulas': {len(aulas_formulas)} fórmulas complejas inyectadas ({locale.upper()})")
    
    return output_file

if __name__ == '__main__':
    # Generar ambas versiones
    generar_excel_con_formulas_fernando(locale='es')  # Español
    print("\n" + "="*80 + "\n")
    generar_excel_con_formulas_fernando(locale='en')  # Inglés
