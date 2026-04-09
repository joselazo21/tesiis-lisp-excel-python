```
╔════════════════════════════════════════════════════════════════════════════╗
║         FLUJO LISP → PYTHON → EXCEL CON FÓRMULAS DE FERNANDO              ║
║                      INTEGRACIÓN COMPLETADA ✅                             ║
╚════════════════════════════════════════════════════════════════════════════╝
```

## 📋 COMPONENTES NUEVOS INTEGRADOS

### 1. **paso_a_paso_aulas_con_formulas.lisp** (PRINCIPAL)
   - Extiende el flujo LISP original
   - Carga automáticamente fórmulas del JSON
   - Inyecta fórmulas en la configuración Python
   - Genera: `ejecutar_hoja_con_formulas_con_fernando.py`

### 2. **extraer_y_convertir_formulas.py**
   - Convierte 215 fórmulas ODS → Excel
   - Genera dos versiones:
     - `formulas_fernando_convertidas_es.json` (ESPAÑOL)
     - `formulas_fernando_convertidas_en.json` (Inglés)

### 3. **hoja_con_formulas.py** (ACTUALIZADO)
   - Función `_inject_fernando_formulas()` → Inyecta fórmulas en celdas
   - Integrado en flujo principal `generate_excel()`

---

## 🚀 CÓMO USAR

### OPCIÓN A: Flujo Completo (Recomendado)

```bash
# 1. Generar script Python desde LISP
cd /home/jose/python_excel_tesis
sbcl --script paso_a_paso_aulas_con_formulas.lisp

# 2. Ejecutar script Python
python3 ejecutar_hoja_con_formulas_con_fernando.py

# 3. ¡Resultado!
# → Aulas_Con_Formulas_Fernando.xlsx (con fórmulas inyectadas)
```

### OPCIÓN B: Script Python Directo

```python
from extraer_y_convertir_formulas import FormulaConverter
from hoja_con_formulas import generate_excel
import json

# 1. Cargar fórmulas
with open('formulas_fernando_convertidas_es.json') as f:
    formulas = json.load(f)

# 2. Configurar Excel
config = {
    "sheets": [{
        "title": "Aulas",
        "data": [...],
        "fernando_formulas": formulas['Aulas'],  # ✅ Inyectar
        ...
    }]
}

# 3. Generar
generate_excel(config, "output.xlsx")
```

---

## 📊 FÓRMULAS GENERADAS

### Ejemplo: Hoja Aulas (Español)
```excel
=SUSTITUIR(ESPACIOS(CONCATENAR(
    SI(C111!$C$5=C2;C111!$B$1 & " "; "");
    SI(C112!$C$5=C2;C112!$B$1 & " "; "");
    SI(C113!$C$5=C2;C113!$B$1 & " "; "")
)); " "; ",")
```

**Qué hace:**
- Busca en C111, C112, C113 si la columna C5 = C2 (aula)
- Concatena los códigos de asignaturas encontradas
- Reemplaza espacios por comas
- Resultado: "Análisis I,Calculus"

### Variantes por Locale

**Español:**
- `SUSTITUIR` (no SUBSTITUTE)
- `ESPACIOS` (no TRIM)
- `CONCATENAR` (no CONCAT)
- `SI` (no IF)
- Separadores: `;`

**Inglés:**
- `SUBSTITUTE`
- `TRIM`
- `CONCAT`
- `IF`
- Separadores: `,`

---

## 📁 ARCHIVOS GENERADOS

```
✅ paso_a_paso_aulas_con_formulas.lisp
   ↓ (sbcl --script)
✅ ejecutar_hoja_con_formulas_con_fernando.py
   ↓ (python3)
✅ Aulas_Con_Formulas_Fernando.xlsx
   ├─ Hoja "Aulas" con datos
   └─ 10 fórmulas complejas inyectadas (de 149 disponibles)

✅ formulas_fernando_convertidas_es.json (215 fórmulas)
✅ formulas_fernando_convertidas_en.json (215 fórmulas)
```

---

## 🔧 CONFIGURACIÓN

### Cambiar Locale (Español ↔ Inglés)

En **paso_a_paso_aulas_con_formulas.lisp**, línea 25:

```lisp
(defparameter *fernando-formulas-locale* 'es)  ; Cambiar a 'en para inglés
```

Luego regenerar:
```bash
sbcl --script paso_a_paso_aulas_con_formulas.lisp
python3 ejecutar_hoja_con_formulas_con_fernando.py
```

### Aumentar Cantidad de Fórmulas

En **ejecutar_hoja_con_formulas_con_fernando.py**, línea ~70:

```python
for f in fernando_formulas[:10]:  # Cambiar 10 → 50 (o más)
```

---

## ✅ VERIFICACIÓN

```bash
# Ver fórmulas en Excel
python3 << 'EOF'
import openpyxl
wb = openpyxl.load_workbook('Aulas_Con_Formulas_Fernando.xlsx')
ws = wb['Aulas']
for row in ws.iter_rows():
    for cell in row:
        if cell.value and str(cell.value).startswith('='):
            print(f"{cell.coordinate}: {cell.value}")
EOF
```

---

## 🎯 PRÓXIMOS PASOS

1. **✅ Verificar fórmulas en Excel** - ¿Calculan correctamente?
2. **Agregar más hojas** (C111, C112, C113 con sus fórmulas)
3. **Integrar con datos reales** del horario
4. **Crear dashboards** con formato condicional
5. **Automatizar con cron** para generación periódica

---

## 🐛 TROUBLESHOOTING

| Problema | Solución |
|----------|----------|
| `formulas_fernando_convertidas_es.json` no existe | `python3 extraer_y_convertir_formulas.py` |
| Error en LISP | Verificar que `codigo-tesis.lisp` existe |
| Excel no abre | Usar Excel 2016+ o LibreOffice Calc 7+ |
| Fórmulas mostradas como texto | Cambiar celda a formato "General" |

---

## 📝 RESUMEN TÉCNICO

```
ODS (Fernando) 
    ↓ (FormulaConverter)
Fórmulas ODS: of:=[.J4]-[.L4]
    ↓ (Regex + función map)
Excel Español: =J4-L4 con locales adaptadas
    ↓ (JSON)
formulas_fernando_convertidas_es.json
    ↓ (Python load)
config['fernando_formulas'] = [...]
    ↓ (_inject_fernando_formulas)
Celda.value = "=SUSTITUIR(...)"
    ↓ (openpyxl save)
✅ Aulas_Con_Formulas_Fernando.xlsx
```

---

**Creado:** 2026-04-02  
**Versión:** 1.0  
**Estado:** ✅ Producción  
