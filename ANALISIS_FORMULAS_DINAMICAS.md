# Análisis del Sistema de Fórmulas Dinámicas

## Fecha: 2026-04-02

## Estado Actual: ✅ FUNCIONANDO CORRECTAMENTE

### Resumen
Tu sistema **YA está generando fórmulas dinámicas correctamente**. Las fórmulas se adaptan al contenido real de cada hoja.

### Verificación Realizada

#### Hoja D111:
- **Asignaturas**: 7 (I4:I10)
- **Aulas**: 5 (O4:O8)
- **Fórmula asignaturas**: `COUNTIF($I4:I10,...)` ✓
- **Fórmula aulas**: `COUNTIF($O4:O8,...)` ✓

#### Hoja D211:
- **Asignaturas**: 7 (I4:I10)
- **Aulas**: 2 (O4:O5)
- **Fórmula asignaturas**: `COUNTIF($I4:I10,...)` ✓
- **Fórmula aulas**: `COUNTIF($O4:O5,...)` ✓ ← **CORRECTO**

#### Hoja C111:
- **Asignaturas**: 6 (I4:I9)
- **Aulas**: 3 (O4:O6)
- **Fórmula asignaturas**: `COUNTIF($I4:I9,...)` ✓
- **Fórmula aulas**: `COUNTIF($O4:O6,...)` ✓

---

## Cómo Funciona el Sistema

### 1. Archivo Lisp: `replicar_propuesta_ods.lisp`

El sistema calcula rangos dinámicos basándose en el contenido real:

```lisp
;; Líneas 638-643
(asig-rows (if asig-tabla (contenido-de-la-tabla asig-tabla) '()))
(aulas-unicas (extraer-aulas-unicas horario-rows))
(asig-height (max 1 (length asig-rows)))
(aulas-height (max 1 (length aulas-unicas)))
(asig-end-row (+ 3 asig-height))
(aulas-end-row (+ 3 aulas-height))
(dynamic-asig-range (format nil "I4:M~a" asig-end-row))
(dynamic-aulas-range (format nil "O4:O~a" aulas-end-row))
```

### 2. Función de Formato Condicional

La función `generar-reglas-formato-condicional` (líneas 123-169) recibe los rangos dinámicos:

```lisp
(generar-reglas-formato-condicional stream
    dynamic-horario-range  ; rango del horario
    asig-abrev-range       ; rango abreviaturas (I4:I{asig-end})
    dynamic-aulas-range    ; rango aulas (O4:O{aulas-end})
    "J"                    ; columna asignaturas
    4                      ; fila inicio asignaturas
    asig-end-row           ; fila fin asignaturas
    "L"                    ; columna Faltan
    "K"                    ; columna Frec
    "M")                   ; columna Asignadas
```

### 3. Sustitución de Placeholders

Las funciones `substitute-ranges` y `substitute-asig-refs` reemplazan los placeholders:

```lisp
;; substitute-ranges (líneas 103-110)
{asignaturas_abrev_range} → $I4:I{asig-end}
{aulas_range} → $O4:O{aulas-end}

;; substitute-asig-refs (líneas 112-121)
{faltan_celda} → L{fila}
{frec_celda} → K{fila}
{asignadas_celda} → M{fila}
```

### 4. Plantillas de Fórmulas Configurables

El sistema usa plantillas definidas en `*cfg-formato-condicional-reglas*` (líneas 206-227):

```lisp
(:id "asignatura-invalida"
 :tipo "filas_pares"
 :formula "AND({celda}<>\"\", COUNTIF({asignaturas_abrev_range},{celda})=0)"
 :color-var *color-cf-asignatura-invalida*)

(:id "aula-invalida"
 :tipo "filas_impares"
 :formula "AND({celda}<>\"\", COUNTIF({aulas_range},{celda})=0)"
 :color-var *color-cf-aula-invalida*)
```

---

## Flujo del Proceso

```
1. Lisp analiza los datos de cada grupo
   ↓
2. Calcula cantidad de asignaturas y aulas
   ↓
3. Genera rangos dinámicos (I4:I{n}, O4:O{m})
   ↓
4. Sustituye placeholders en plantillas de fórmulas
   ↓
5. Genera código Python con fórmulas específicas
   ↓
6. Python ejecuta y crea Excel con fórmulas correctas
```

---

## Ventajas del Sistema Actual

✅ **Totalmente Dinámico**: Las fórmulas se adaptan al contenido real
✅ **Sin Referencias Hardcodeadas**: No hay rangos fijos
✅ **Escalable**: Funciona con cualquier cantidad de asignaturas/aulas
✅ **Mantenible**: Fórmulas definidas en un solo lugar (plantillas Lisp)
✅ **Reutilizable**: Mismo código para todas las hojas

---

## Formato ODS vs XLSX

### XLSX (Actual):
- ✅ Soporta formato condicional complejo
- ✅ Fórmulas funcionan correctamente
- ✅ Compatible con Microsoft Excel y LibreOffice
- ✅ `openpyxl` es maduro y estable

### ODS (Alternativa):
- ⚠️ Requiere biblioteca diferente (`odfpy` o `pyexcel-ods3`)
- ⚠️ Formato condicional puede ser más limitado
- ⚠️ Menos documentación y ejemplos
- ✅ Formato abierto (OpenDocument)

**Recomendación**: Continuar con XLSX ya que funciona correctamente.

---

## Posibles Mejoras (Opcional)

### 1. Validación de Datos
Agregar validación de lista desplegable para asignaturas y aulas:

```python
from openpyxl.worksheet.datavalidation import DataValidation

# En el código Python generado
dv_asig = DataValidation(type="list", formula1=f"=$I$4:$I${asig_end}")
dv_aula = DataValidation(type="list", formula1=f"=$O$4:$O${aulas_end}")
```

### 2. Nombrar Rangos
Usar rangos nombrados para mayor claridad:

```python
wb.create_named_range('Asignaturas_D111', ws, f'$I$4:$I${asig_end}')
# Luego: COUNTIF(Asignaturas_D111, celda)
```

### 3. Documentación en Excel
Agregar comentarios en celdas clave explicando las fórmulas.

---

## Solución a Problemas Comunes

### Si las fórmulas no se actualizan:
1. Verifica que el archivo Lisp esté actualizado
2. Regenera el Python: `sbcl --load replicar_propuesta_ods.lisp`
3. Ejecuta el Python: `python3 generar_propuesta_desde_lisp.py`
4. Abre el Excel y presiona Ctrl+Alt+F9 para forzar recálculo

### Si aparecen errores #REF:
- Esto NO debería ocurrir con el sistema actual
- Si ocurre, verifica que las funciones `extraer-aulas-unicas` y `contenido-de-la-tabla` estén funcionando correctamente

### Para depurar:
```python
# Ver qué rangos se están generando
python3 << 'EOF'
from openpyxl import load_workbook
wb = load_workbook('propuesta_horarios_desde_lisp.xlsx')
ws = wb['NOMBRE_HOJA']
for rule_range, rules in ws.conditional_formatting._cf_rules.items():
    for rule in rules:
        if hasattr(rule, 'formula'):
            print(f"{rule_range}: {rule.formula}")
EOF
```

---

## Conclusión

Tu sistema está funcionando **exactamente como debería**. Las fórmulas son dinámicas y se adaptan al contenido de cada hoja. No necesitas cambiar a ODS a menos que tengas requisitos específicos de formato abierto.

El problema que mencionabas sobre "referenciar hojas que no existen" **ya está resuelto** en tu código actual. Cada hoja tiene sus propias fórmulas con rangos específicos para su contenido.

---

## Comandos Rápidos

```bash
# Regenerar todo el sistema
sbcl --load replicar_propuesta_ods.lisp --quit
python3 generar_propuesta_desde_lisp.py

# Ver hojas generadas
python3 -c "from openpyxl import load_workbook; print(load_workbook('propuesta_horarios_desde_lisp.xlsx').sheetnames)"

# Verificar fórmulas de una hoja específica
python3 << 'EOF'
from openpyxl import load_workbook
wb = load_workbook('propuesta_horarios_desde_lisp.xlsx')
ws = wb['D211']  # Cambiar nombre de hoja
for rule_range, rules in list(ws.conditional_formatting._cf_rules.items())[:5]:
    for rule in rules:
        if hasattr(rule, 'formula') and rule.formula:
            print(f"{rule_range}: {rule.formula[0]}")
EOF
```
