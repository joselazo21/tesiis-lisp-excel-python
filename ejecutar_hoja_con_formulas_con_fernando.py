import sys
import json
from hoja_con_formulas import generar_excel_personalizado


import json

def cargar_formulas_fernando(locale='es'):
    '''Carga fórmulas convertidas del ODS de Fernando'''
    filename = f'formulas_fernando_convertidas_{locale}.json'
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f'⚠️  {filename} no encontrado')
        return {}

# Cargar fórmulas
_formulas_fernando = cargar_formulas_fernando('es')


data_aulas = []
headers_aulas = ["Lunes", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"]

data_aulas.append(["1ro", "", "", "", "", "", "C111,C112", "", "", "", "C113"])
data_aulas.append(["2do", "", "", "", "", "C112", "C113", "", "", "", ""])
data_aulas.append(["3ro", "", "", "", "", "", "C111,C113", "", "C112", "", ""])
data_aulas.append(["4to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["5to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["6to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append([])
data_aulas.append(["Martes", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"])
data_aulas.append(["1ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["2do", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["3ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["4to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["5to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["6to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append([])
data_aulas.append(["Miércoles", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"])
data_aulas.append(["1ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["2do", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["3ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["4to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["5to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["6to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append([])
data_aulas.append(["Jueves", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"])
data_aulas.append(["1ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["2do", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["3ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["4to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["5to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["6to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append([])
data_aulas.append(["Viernes", "Aula 1", "Aula 2", "Aula 3", "Aula 4", "Aula 5", "Aula 6", "Aula 7", "Aula 8", "Aula 9", "Lab"])
data_aulas.append(["1ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["2do", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["3ro", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["4to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["5to", "", "", "", "", "", "", "", "", "", ""])
data_aulas.append(["6to", "", "", "", "", "", "", "", "", "", ""])

# ============================================================
# INYECTAR FÓRMULAS COMPLEJAS DEL ODS DE FERNANDO
# ============================================================

fernando_formulas = _formulas_fernando.get('Aulas', [])

# Convertir fórmulas a formato esperado
formulas_para_inyectar = []
for f in fernando_formulas[:10]:  # Primeras 10 de demostración
    formulas_para_inyectar.append({
        'cell': f.get('cell'),
        'formula': f.get('excel')
    })

config_excel = {
    'sheets': [
        {
            'title': 'Aulas',
            'headers': headers_aulas,
            'data': data_aulas,
            'fernando_formulas': formulas_para_inyectar,  # ✅ NUEVO
            'column_widths': {i: 12 for i in range(1, 12)},
            'header_style': {'bold': True, 'align': 'center', 'bg_color': 'F09E9E'}
        }
    ]
}

# Generar Excel con fórmulas inyectadas
print('✅ Generando Excel con fórmulas de Fernando...')
generar_excel_personalizado(config_excel, 'Aulas_Con_Formulas_Fernando.xlsx')
print(f'✅ Archivo generado: Aulas_Con_Formulas_Fernando.xlsx')
print(f'✅ Total de fórmulas inyectadas: {len(formulas_para_inyectar)}')
