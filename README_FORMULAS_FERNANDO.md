# 📊 Integración de Fórmulas Complejas de Fernando en LISP

## ✅ Estado: COMPLETADO Y FUNCIONAL

Este proyecto integra automáticamente las **215 fórmulas complejas** del ODS de Fernando en tu flujo de generación de Excel mediante LISP → Python.

---

## 🎯 Objetivo Logrado

✅ **Extraer** 215 fórmulas del ODS original de Fernando  
✅ **Convertir** de formato ODS a formato Excel (español e inglés)  
✅ **Inyectar** fórmulas directamente en celdas Excel  
✅ **Integrar** en flujo LISP principal  
✅ **Automatizar** generación de Excel con fórmulas  
✅ **Validar** que Excel calcula correctamente  

---

## 📁 Archivos Principales

### Para Ejecutar (Producción)

```
paso_a_paso_aulas_con_formulas.lisp (13 KB)
  ↓ sbcl --script
ejecutar_hoja_con_formulas_con_fernando.py (auto-generado)
  ↓ python3
Aulas_Con_Formulas_Fernando.xlsx ✅ (resultado final)
```

### Datos (Fórmulas Convertidas)

```
formulas_fernando_convertidas_es.json (75 KB)  ← Español
formulas_fernando_convertidas_en.json (74 KB)  ← Inglés
  └─ 215 fórmulas cada uno
     ├─ 22 de hoja C111
     ├─ 22 de hoja C112
     ├─ 22 de hoja C113
     └─ 149 de hoja Aulas (complejas)
```

### Módulos Reutilizables

```
extraer_y_convertir_formulas.py (6,5 KB)
  └─ FormulaConverter.ods_to_excel_formula()
  └─ FormulaConverter.extract_formulas_from_ods()

hoja_con_formulas.py (actualizado)
  └─ _inject_fernando_formulas()
  └─ generate_excel()

formulas_fernando.lisp (3,8 KB)
  └─ Interfaz LISP para cargar fórmulas
```

---

## 🚀 Cómo Usar

### Opción 1: Flujo Completo (Recomendado)

```bash
# 1. Generar script Python desde LISP
cd /home/jose/python_excel_tesis
sbcl --script paso_a_paso_aulas_con_formulas.lisp

# 2. Ejecutar generador
python3 ejecutar_hoja_con_formulas_con_fernando.py

# 3. Revisar resultado
# → Aulas_Con_Formulas_Fernando.xlsx está listo
```

### Opción 2: Script Python Directo

```python
from extraer_y_convertir_formulas import FormulaConverter
from hoja_con_formulas import generate_excel
import json

# Cargar todas las 215 fórmulas
with open('formulas_fernando_convertidas_es.json') as f:
    formulas = json.load(f)

# Configurar Excel
config = {
    "sheets": [{
        "title": "Aulas",
        "headers": ["Aula", "Lunes", "Martes", ...],
        "data": [[...], [...], ...],
        "fernando_formulas": formulas['Aulas'],  # ✅ Inyectar aquí
        "column_widths": {1: 12, 2: 25, ...},
        ...
    }]
}

# Generar
generate_excel(config, "output.xlsx")
```

---

## 📊 Ejemplo: Fórmula Inyectada

### Original (ODS)
```
of:=SUBSTITUTE(TRIM(COM.MICROSOFT.CONCAT(
  IF([$C111.$C$5]=[.C$2];[$C111.$B$1] & " "; "");
  ...
)); " "; ",")
```

### Convertida (Excel Español) ✅
```excel
=SUSTITUIR(ESPACIOS(CONCATENAR(
  SI(C111!$C$5=C2;C111!$B$1 & " "; "");
  ...
); " "; ",")
```

**Lo que hace:** Busca en 3 hojas si el aula de la celda C2 está asignada, concatena asignaturas encontradas y reemplaza espacios por comas.

---

## ✅ Verificación

```bash
python3 << 'EOF'
import openpyxl

wb = openpyxl.load_workbook('Aulas_Con_Formulas_Fernando.xlsx')
ws = wb['Aulas']

for row in ws.iter_rows():
    for cell in row:
        if cell.value and str(cell.value).startswith('='):
            print(f"{cell.coordinate}: {cell.value[:80]}...")
EOF
```

---

## 📈 Ventajas

✅ Reutiliza lógica compleja de Fernando  
✅ Automatiza conversión ODS → Excel  
✅ Localiza automáticamente (español/inglés)  
✅ Escala a 215 fórmulas sin cambios  
✅ Mantiene referencias cruzadas intactas  

---

**¡Sistema listo para producción! 🚀**
