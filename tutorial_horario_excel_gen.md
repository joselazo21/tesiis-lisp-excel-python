# Tutorial: Construir un Horario con excel_gen.py

Este tutorial te muestra paso a paso cómo crear un archivo Excel con horarios de grupos usando `excel_gen.py` y `generador_formulas.py`.

---

## Paso 1: Imports y Configuración Básica

```python
from excel_gen import generate_excel, crear_hoja_grupo_vacia, crear_hoja_aulas, aulas_formulas_generator
from generador_formulas import build_aulas_fernando_formulas
```

**Configuración del horario:**
```python
GRUPOS = ["D111", "D211", "C111"]  # ← Grupos que quieres crear
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
AULAS = [f"Aula {i}" for i in range(1, 10)] + ["Lab"]
TURNOS = 6
ROW_STEP = 3  # filas por turno (nombre + aula + separador)
```

---

## Paso 2: Crear Hojas de Grupos

Cada grupo tiene su propia hoja con:
- Horario vacío (columnas para cada día)
- Tabla de asignaturas vacía
- Lista de aulas

```python
# Crear una hoja por cada grupo
sheets = [crear_hoja_grupo_vacia(g) for g in GRUPOS]
```

---

## Paso 3: Crear Hoja Aulas

La hoja Aulas muestra la tabla cruzada de qué grupo está en qué aula cada día/turno:

```python
# Crear estructura vacía de la hoja Aulas
aulas_cfg = crear_hoja_aulas(GRUPOS)

# Añadir el generador de fórmulas cruzadas (grupos → aulas)
aulas_cfg['formulas_generator'] = aulas_formulas_generator(GRUPOS)

# Agregar al listado de hojas
sheets.append(aulas_cfg)
```

El `formulas_generator` es una función que yield diccionarios con:
- `cell`: referencia Excel (ej: "C4")
- `formula`: la fórmula (puede usar referencias cruzadas entre hojas)

---

## Paso 4: Generar el Excel

```python
config = {"sheets": sheets}
generate_excel(config, "mi_horario.xlsx")
```

---

## Estructura de una Hoja de Grupo

El diccionario de cada hoja tiene estas claves:

| Clave | Descripción |
|-------|-------------|
| `title` | Nombre de la hoja |
| `data` | Filas de datos (listas) |
| `column_widths` | Ancho de columnas {col: ancho} |
| `range_styles` | Colores por rango [{"range": "B3:L3", "style": {"bg_color": "E6B8AF"}}] |
| `table_ranges` | Rangos para bordes |
| `table_block_sizes` | Bordes por bloques {"range": "B4:B15", "row_step": 3} |
| `merge_ranges` | Celdas a fusionar ["B4:B6"] |
| `table_borders` | True para activar bordes |
| `formulas` | Fórmulas simples [{"row": 5, "col": 13, "value": "=COUNTIF(...)"}] |
| `conditional_format_rules` | Formato condicional avanzado |

---

## Ejemplo Completo

Guarda esto como `generar_mi_horario.py`:

```python
from excel_gen import generate_excel, crear_hoja_grupo_vacia, crear_hoja_aulas

# === Configuración ===
GRUPOS = ["D111", "D211", "C111"]
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
AULAS = [f"Aula {i}" for i in range(1, 10)] + ["Lab"]
TURNOS = 6
ROW_STEP = 3

# === Generar hojas ===
sheets = [crear_hoja_grupo_vacia(g) for g in GRUPOS]

# Hoja Aulas con fórmulas cruzadas
aulas_cfg = crear_hoja_aulas(GRUPOS)
aulas_cfg['formulas_generator'] = lambda: (
    formula for formula in build_aulas_fernando_formulas(GRUPOS, DIAS, ROW_STEP, TURNOS)
) if GRUPOS else iter([])

# Ejecutar
from generador_formulas import build_aulas_fernando_formulas
aulas_cfg['formulas_generator'] = lambda: build_aulas_fernando_formulas(GRUPOS, DIAS, ROW_STEP, TURNOS)
sheets.append(aulas_cfg)

generate_excel({"sheets": sheets}, "mi_horario.xlsx")
```

Ejecuta:
```bash
python generar_mi_horario.py
```

---

## Personalizar una Hoja

Si necesitas una hoja personalizada, crea el diccionario directamente:

```python
mi_hoja = {
    "title": "MiGrupo",
    "data": [
        ["Grupo", "MiGrupo"],
        ["", "", "Lunes", "Martes", "Miércoles"],
        ["Turno 1", "", "Aula 1", ""],
        ["", "", "Aula 2", "Aula 3"],
    ],
    "column_widths": {1: 14, 2: 14, 3: 20},
    "table_borders": True,
    "range_styles": [
        {"range": "A3:D3", "style": {"bold": True, "bg_color": "E6B8AF"}},
    ],
    "merge_ranges": ["A4:A5"],
}
```

---

##keys a usar en conditional_format_rules

| Tipo | Descripción |
|------|-------------|
| `rango` | Aplica a cada celda |
| `filas_pares` | Filas 1,3,5... relativas |
| `filas_impares` | Filas 2,4,6... relativas |
| `pares_con_siguiente` | Evalúa par + siguiente |

Placeholders: `{celda}`, `{celda_siguiente}`, `{fila}`, `{columna}`

---

##keys para fórmulas complejas (formulas_generator)

```python
def mis_formulas():
    yield {"cell": "C4", "formula": "=D111!C4"}
    yield {"cell": "D5", "formula": "=IF(D111!C5<>\"\", D111!C5, \"\")"}
```

Cada yield devuelve un dict con `cell` y `formula` (o `excel`).

El motor automáticamente:
- Convierte `SI()` → `IF()`
- Convierte `SUSTITUIR()` → `SUBSTITUTE()`
- Convierte `CONCATENAR(a;b;c)` → `(a & b & c)`
- Convierte `;` → `,`
- Convierte referencias `[Sheet1.Cell]` → `Sheet1!Cell`