# Informe Técnico del Sistema de Generación de Horarios

## Resumen Ejecutivo

Este proyecto constituye un sistema completo de **generación automatizada de horarios** desarrollado como parte de una investigación de tesis. El sistema es capaz de:

1. **Generar horarios académicos universitarios** (facultad MATCOM) en formato Excel con fórmulas dinámicas
2. **Generar horarios de programación televisiva** con cálculo automático de tiempos
3. **Transformar datos JSON en archivos Excel** con formato profesional, fórmulas, bordes, colores y formato condicional
4. **Utilizar Lisp como generador de código Python** (metaprogramación)

El sistema puede crear:
- **18 hojas Excel** (17 grupos académicos + hoja "Aulas" consolidada)
- **140+ fórmulas dinámicas** que cruzan entre hojas
- **Formulas deTV** con cálculo automático de horas de inicio/fin
- **Formato condicional** para validación de datos
- **Estilos y bordes** profesionales

---

## 1. Arquitectura General del Sistema

El sistema opera mediante un **pipeline de dos lenguajes**:

```
┌─────────────────┐      ┌─────────────────┐      ┌─────────────────┐
│     JSON        │ ──►  │     LISP        │ ──►  │    PYTHON      │
│  (datos fuente) │      │ (generador)     │      │ (motor Excel)   │
└─────────────────┘      └─────────────────┘      └─────────────────┘
```

**Flujo principal:**
1. **JSON** → contiene los datos crudos de horarios
2. **convertir_json_a_lisp.py** → transforma JSON a variables Lisp
3. **SBCL (Lisp)** → genera código Python declarativo
4. **Python + openpyxl** → crea el archivo Excel final

---

## 2. Scripts Python del Sistema

### 2.1. `convertir_json_a_lisp.py`

**Propósito:** Convierte archivos JSON con datos de horarios en variables Lisp.

**INPUT (horario.json):**
```json
{
  "horarios": [
    {
      "carrera": "Ciencia de Datos",
      "grupo": "D111",
      "tabla": [
        {"Turno": "1", "Lunes": "F", "Martes": "ICD", "Miércoles": "AL"},
        {"Turno": "1", "Lunes": "Aula 8", "Martes": "Aula 7", "Miércoles": "Aula 7*"}
      ]
    }
  ]
}
```

**CÓDIGO que lo procesa:**
```python
def procesar_grupo_nuevo(grupo_data):
    """Procesa un grupo del nuevo formato JSON."""
    tabla = grupo_data.get("tabla", [])
    horario_matriz = [[("", "") for _ in range(5)] for _ in range(6)]
    
    for fila in tabla:
        turno_num = int(fila.get("Turno", "1")) - 1
        for dia_idx, dia_key in enumerate(dias_keys):
            valor = fila.get(dia_key, "")
            if valor:
                asig, aula = parsear_entrada(valor)
                horario_matriz[turno_num][dia_idx] = (asig, aula)
    
    # Convierte a formato Lisp (lista de listas)
    return horario_matriz
```

**OUTPUT (variables_horario.lisp):**
```lisp
;; CIENCIA DE DATOS - D111
(defparameter *horario-d111*
  '(
   ("F" "ICD" "AL" "AM I" "AL")
   ("Aula 8" "Aula 7" "Aula 7*" "Aula 7*" "Aula 7*")
   ("L" "AL" "EF" "L" "")
   ("Aula 6*" "Aula 6*" "SEDER" "Aula 7" "")
   ...
  ))
```

---

### 2.2. `excel_gen.py` / `hoja_con_formulas.py`

**Propósito:** Framework principal de generación de Excel. Es el "motor" que crea archivos .xlsx con formato profesional.

**INPUT (configuración declarativa):**
```python
config = {
    "sheets": [
        {
            "title": "D111",
            "data": [
                ["Grupo", "D111", "", "", ""],
                ["", "", "", "", ""],
                ["", "Turno", "Lunes", "Martes", "Miércoles"],
                ["1", "", "AL", "AM I", "F"],
                ["", "Aula 8", "Aula 7", "Aula 6", ""]
            ],
            "formulas": [
                {"row": 11, "col": 13, "value": "=COUNTA(I4:I10)"},
                {"row": 4, "col": 13, "value": "=COUNTIF($C$4:$G$15,I4)"}
            ],
            "table_borders": True,
            "border_color": "4F81BD",
            "conditional_format_rules": [...]
        }
    ]
}
```

**CÓDIGO que lo procesa:**
```python
def generate_excel(config: dict, filename: str) -> None:
    """Genera un archivo .xlsx completo desde configuración declarativa."""
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_cfg in config.get("sheets", []):
        ws = wb.create_sheet(title=sheet_cfg["title"])
        
        # Pipeline de procesamiento (11 etapas):
        _process_headers(ws, sheet_cfg)      # Encabezados
        _process_data(ws, sheet_cfg)         # Filas de datos
        _process_formulas(ws, sheet_cfg)    # Fórmulas simples
        _inject_fernando_formulas(ws, sheet_cfg)  # Fórmulas cruzadas
        _process_column_widths(ws, sheet_cfg)
        _process_table_borders(ws, sheet_cfg)
        _process_range_styles(ws, sheet_cfg)
        _process_merge_ranges(ws, sheet_cfg)
        _process_conditional_format_rules(ws, sheet_cfg)
    
    wb.save(filename)
```

**OUTPUT (Excel):**
- Archivo `.xlsx` con múltiples hojas
- Fórmulas funcionales
- Bordes, colores, formato condicional

---

### 2.3. `generar_horario_tv_desde_lisp.py`

**Propósito:** Genera horarios de programación televisiva en Excel.

**INPUT:**
```python
planificacion_semanal = [
    {'dia': "lunes", 'programas': [
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "17:30", 
         'hora_final': "18:00", 'tipo_programa': "informativo", 'tipo_publico': "adulto"},
        {'nombre': "POWER RANGERS", 'duracion': 30, 'hora_inicio': "18:00", 
         'hora_final': "18:30", 'tipo_programa': "infantil", 'tipo_publico': "infantil"},
        ...
    ]},
    {'dia': "martes", 'programas': [...]},
    ...
]
```

**CÓDIGO que lo procesa (en hoja_con_formulas.py):**
```python
def construir_hoja_tv_desde_parametros(dia_cfg, index, start_time_param_cell):
    """Construye configuración para una hoja de día de TV."""
    
    formulas = []
    for idx in range(len(planificacion)):
        row = 4 + idx
        # Fórmula: hora de inicio = hora final del programa anterior
        formulas.append({
            "row": row,
            "col": 5,  # columna E
            "value": f'=IF(G{row}="","",F{row - 1})'
        })
        # Fórmula: hora final = inicio + duración
        formulas.append({
            "row": row,
            "col": 6,  # columna F
            "value": f'=IF(G{row}="","",TEXT(TIMEVALUE(E{row}) + IFERROR(VLOOKUP(G{row},$A$4:$C${prog_end_row},2,FALSE),0)/1440,"hh:mm"))'
        })
    
    return {
        "title": title,
        "data": data,
        "formulas": formulas,
        "table_borders": True,
        "conditional_format_rules": [...]
    }
```

**OUTPUT (Excel):**
- 7 hojas (una por día) + hoja Resumen
- Columnas: Programa, Duración, Tipo, Hora inicio, Hora final
- Fórmulas que calculan automáticamente tiempos
- Formato condicional que detecta errores de programación

---

### 2.4. `generador_formulas.py`

**Propósito:** Generador de fórmulas cruzadas para la hoja "Aulas".

**INPUT:**
```python
groups = ["C111", "C121", "C211", "D111", "M111", ...]  # 17 grupos
dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
```

**CÓDIGO:**
```python
def build_aulas_formula(groups, group_cell_ref, header_ref):
    """Fórmula que concatena grupos coincidentes para una celda."""
    parts = [
        f'IF({group}!{group_cell_ref}={header_ref},{group}!$B$1 & " ","")'
        for group in groups
    ]
    return f'=SUBSTITUTE(TRIM(CONCAT({",".join(parts)}))," ",",")'

def build_aulas_fernando_formulas(groups, dias, row_step=3, turnos=6):
    """Genera fórmulas cruzadas para cada celda de la hoja Aulas."""
    formulas = []
    for header_row, row_start, row_end, group_col in build_day_blocks(dias, turnos):
        for aulas_col in range(3, 13):  # Columnas C a L (aulas)
            for aulas_row in range(row_start, row_end + 1):
                formulas.append({
                    "cell": f"{aulas_col_letter}{aulas_row}",
                    "formula": build_aulas_formula(groups, group_cell_ref, header_ref)
                })
    return formulas
```

**OUTPUT:**
```python
[
    {"cell": "C4", "formula": "=SUBSTITUTE(TRIM(CONCAT(IF(C111!$C$4=$C$2,C111!$B$1 & \" \",\"\") ...)),\" \",\",\")"},
    {"cell": "C5", "formula": "=SUBSTITUTE(TRIM(CONCAT(IF(C111!$C$5=$C$2,...)))"},
    ...
]
```

**Fórmula Excel resultante (ejemplo para celda F4):**
```excel
=SUSTITUIR(ESPACIOS(CONCATENAR(
  SI(D111!$C$5=F$2;D111!$B$1 & " "; "")
  SI(C111!$C$5=F$2;C111!$B$1 & " "; "")
  SI(C121!$C$5=F$2;C121!$B$1 & " "; "")
  ...
)); " "; ",")
```

---

## 3. Sistema de Generación desde LISP

### 3.1. Flujo Principal

```
horario.json → convertir_json_a_lisp.py → variables_horario.lisp
                                                    │
                                                    ▼
                                              replicar_propuesta_ods.lisp (SBCL)
                                                    │
                                                    ▼
                                         generar_propuesta_desde_lisp.py
                                                    │
                                                    ▼
                                         propuesta_horarios_desde_lisp.xlsx
```

### 3.2. Archivos LISP principales

#### `codigo-tesis.lisp` - Framework base
Define las clases fundamentales del sistema:

```lisp
(defclass clase-tabla ()
  ((filas :initarg :filas :accessor filas)
   (columnas :initarg :columnas :accessor columnas)
   (contenido :initarg :contenido :accessor contenido)
   (estilos :initarg :estilos :accessor estilos)))

(defclass clase-hoja ()
  ((nombre :initarg :nombre :accessor nombre)
   (horario :initarg :horario :accessor horario)  ; clase-tabla
   (asignaturas :initarg :asignaturas :accessor asignaturas)))

(defclass clase-libro ()
  ((hojas :initarg :hojas :accessor hojas)))
```

#### `variables_horario.lisp` - Datos
Contiene los horarios de los 17 grupos:

```lisp
;; 17 grupos definidos como defparameter
(defparameter *horario-d111* '(...))  ; Ciencia de Datos 1er año
(defparameter *horario-d211* '(...))  ; Ciencia de Datos 2do año
(defparameter *horario-c111* '(...))  ; Ciencias de la Computación 1er año
(defparameter *horario-m111* '(...))  ; Matemática 1er año
;; ... hasta 17 grupos
```

#### `replicar_propuesta_ods.lisp` - Orquestador

**INPUT (Lisp):**
```lisp
(load "codigo-tesis.lisp")
(load "variables_horario.lisp")

;; Crea objeto para grupo D111
(crear-hoja-grupo "D111" *horario-d111* *asignaturas-d111*)
```

**CÓDIGO (Lisp) que procesa:**
```lisp
(defun crear-hoja-grupo (nombre-grupo datos-horario datos-asignaturas)
  "Crea una hoja de grupo con estructura completa."
  (let ((hoja (make-instance 'clase-hoja :nombre nombre-grupo)))
    ;; Genera estructura de datos para Excel
    (setf (horario hoja) (crear-tabla-horario datos-horario))
    (setf (asignaturas hoja) (crear-tabla-asignaturas datos-asignaturas))
    ;; Genera código Python
    (generar-codigo-python hoja)
    hoja))
```

**OUTPUT (Python generado):**
```python
sheets_cfg.append({
    "title": "D111",
    "data": [
        ["Grupo ", "D111", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["", "Turno", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes",
         "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"],
        ["1", "", "AL", "AM I", "F", "", "", "AL", "Álgebra Lineal", 3, 3, 0, "", "Aula 8"],
        ...
    ],
    "formulas": [
        {"row": 11, "col": 12, "value": "Total:"},
        {"row": 11, "col": 13, "value": "=COUNTA(I4:I10)"},
    ],
    "conditional_format_rules": [...],
    "merge_ranges": ["B4:B5", "B6:B7"],
    "range_styles": [{"range": "I3:I10", "style": {"bg_color": "A9D18E"}}],
    "table_borders": True
})
```

---

### 3.3. Sistema Integrado (`sistema_horarios_integrado.lisp`)

Este archivo integra todo el sistema con detección automática de grupos:

```lisp
(defun obtener-todos-los-grupos ()
  "Obtiene todos los grupos definidos automáticamente."
  (let ((grupos '()))
    (dolist (sym (list-all-packages))
      (do-symbols (s sym)
        (when (and (boundp s)
                   (string-prefix-p "*HORARIO-" (string s)))
          (let* ((sym-name (string s))
                 (grupo-name (subseq sym-name (length "*HORARIO-")
                                     (1- (length sym-name)))))
            (pushnew (intern grupo-name) grupos)))))
    (sort grupos #'string<)))
```

**OUTPUT:**
```lisp
(C111 C121 C211 C311 C411 D111 D211 D311 D411 M111 M211 M311 M411)
```

---

## 4. Ejemplos Completos de Input → Proceso → Output

### Ejemplo 1: Generación de Horario Académico

**INPUT (horario.json fragmento):**
```json
{
  "horarios": [
    {
      "carrera": "Ciencia de Datos",
      "grupo": "D111",
      "tabla": [
        {"Turno": "1", "Lunes": "F", "Martes": "ICD", "Miércoles": "AL", "Jueves": "AM I", "Viernes": "AL"},
        {"Turno": "1", "Lunes": "Aula 8", "Martes": "Aula 7", "Miércoles": "Aula 7*", "Jueves": "Aula 7*", "Viernes": "Aula 7*"}
      ]
    }
  ]
}
```

**PROCESO:**
```bash
# Paso 1: Convertir JSON a Lisp
python3 convertir_json_a_lisp.py

# Paso 2: Generar código Python desde Lisp
sbcl --script replicar_propuesta_ods.lisp

# Paso 3: Generar Excel desde Python
python3 generar_propuesta_desde_lisp.py
```

**OUTPUT (propuesta_horarios_desde_lisp.xlsx):**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│                          Grupo D111                                         │
├────┬────────┬─────────┬─────────┬───────────┬─────────┬─────────┬───────────┤
│    │ Turno  │ Lunes   │ Martes  │ Miércoles │ Jueves  │ Viernes │           │
├────┼────────┼─────────┼─────────┼───────────┼─────────┼─────────┼───────────┤
│    │Turno 1 │   AL    │   AM I  │    F      │   ICD   │   L     │           │
│    │        │ Aula 8  │ Aula 7  │  Aula 6   │ Aula 7  │ Aula 6  │           │
├────┼────────┼─────────┼─────────┼───────────┼─────────┼─────────┼───────────┤
│    │Turno 2 │   L     │   AL    │    EF     │   L     │         │           │
│    │        │ Aula 6* │Aula 6*  │  SEDER    │Aula 7*  │         │           │
├────┼────────┴─────────┴─────────┴───────────┴─────────┴─────────┼───────────┤
│    │ Abrev │ Asignaturas      │ Frec │ Faltan │ Asignadas│ Aulas  │
├────┼───────┼──────────────────┼──────┼────────┼──────────┼────────┤
│    │  AL   │ Álgebra Lineal   │  3   │ =K4-M4 │=COUNTIF  │ ...    │
│    │  L    │ Lógica           │  2   │ =K5-M5 │=COUNTIF  │ ...    │
│    │  F    │ Fund. Programación│ 3   │ =K6-M6 │=COUNTIF  │ ...    │
├────┼───────┼──────────────────┼──────┼────────┼──────────┼────────┤
│    │       │                  │ Σ=8  │        │   8      │        │
└────┴───────┴──────────────────┴──────┴────────┴──────────┴────────┘
```

---

### Ejemplo 2: Generación de Fórmulas Cruzadas (Hoja Aulas)

**INPUT (grupos detectados):**
```lisp
(obtener-todos-los-grupos)
;; → (C111 C121 C211 C311 C411 D111 D211 D311 D411 M111 M211 M311 M411)
```

**PROCESO (generador_formulas.py):**
```python
groups = ["C111", "C121", "C211", "C311", "C411", "D111", "D211", "D311", "D411", "M111", "M211", "M311", "M411"]
formulas = build_aulas_fernando_formulas(groups, dias=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"])
```

**OUTPUT (fórmulas generadas, ejemplo para celda F4):**
```python
{
    "cell": "F4",
    "formula": "=SUBSTITUTE(TRIM(CONCAT(IF(C111!$C$5=$F$2,C111!$B$1 & \" \",\"\") ,IF(C121!$C$5=$F$2,C121!$B$1 & \" \",\"\") ,IF(C211!$C$5=$F$2,C211!$B$1 & \" \",\"\") ,IF(D111!$C$5=$F$2,D111!$B$1 & \" \",\"\") ,IF(M111!$C$5=$F$2,M111!$B$1 & \" \",\"\"))),\" \",\",\")"
}
```

**OUTPUT (Excel resultante):**
```
┌─────────┬──────────┬──────────┬──────────┬──────────┬──────────┐
│         │Aula 1    │Aula 2   │Aula 3    │Aula 4    │Aula 5    │
├─────────┼──────────┼──────────┼──────────┼──────────┼──────────┤
│ LUNES   │          │          │          │          │          │
│  8:00   │ D111     │ C112     │          │          │          │
│ 10:00   │          │          │ C211     │          │ D111     │
├─────────┼──────────┼──────────┼──────────┼──────────┼──────────┤
│ MARTES  │          │          │          │          │          │
│  8:00   │ C111     │          │ D311     │          │          │
└─────────┴──────────┴──────────┴──────────┴──────────┴──────────┘
```

Cada celda muestra qué grupos están en esa aula a esa hora.

---

### Ejemplo 3: Generación de Horario de TV

**INPUT (generar_horario_tv_desde_lisp.py):**
```python
planificacion_semanal = [
    {'dia': "lunes", 'programas': [
        {'nombre': "HABANA NOTICIARIO", 'duracion': 30, 'hora_inicio': "17:30", 
         'hora_final': "18:00", 'tipo_programa': "informativo"},
        {'nombre': "POWER RANGERS", 'duracion': 30, 'hora_inicio': "18:00", 
         'hora_final': "18:30", 'tipo_programa': "infantil"},
    ]},
    {'dia': "martes", 'programas': [...]}
]
```

**PROCESO:**
```bash
python3 generar_horario_tv_desde_lisp.py
```

**OUTPUT (horario_tv_semanal.xlsx):**
```
┌─────────────────────────────────────────────────────────────────────┐
│ Programacion Lunes                                                  │
├─────────┬──────────────┬────────┬────────────┬────────────────────┤
│Programa │ Duracion(min)│  Tipo  │Hora inicio │ Hora terminacion   │
├─────────┼──────────────┼────────┼────────────┼────────────────────┤
│HABANA   │      30      │inform. │  17:30     │    18:00           │
│NOTICIARIO│             │        │=IF(G4="",F3)│=TEXT(TIMEVALUE... │
├─────────┼──────────────┼────────┼────────────┼────────────────────┤
│POWER    │      30      │infantil│  18:00     │    18:30           │
│RANGERS  │              │        │=IF(G5="",F4)│=TEXT(TIMEVALUE... │
└─────────┴──────────────┴────────┴────────────┴────────────────────┘
```

---

## 5. Comandos de Ejecución

### Ejecución rápida:
```bash
./ru.sh
```

### Paso a paso:
```bash
# 1. Convertir JSON a Lisp
python3 convertir_json_a_lisp.py

# 2. Generar código Python desde LISP
sbcl --script replicar_propuesta_ods.lisp

# 3. Generar Excel
python3 generar_propuesta_desde_lisp.py
```

### Para horario de TV:
```bash
python3 generar_horario_tv_desde_lisp.py
```

---

## 6. Características Técnicas del Sistema

### Fórmulas Soportadas:
- **Fórmulas simples:** `=COUNTA()`, `=SUM()`, `=COUNTIF()`, `=K4-M4`
- **Fórmulas cruzadas:** Referencias entre hojas (`C111!$C$5=F$2`)
- **Fórmulas con funciones anidadas:** `SUBSTITUTE(TRIM(CONCAT(...)))`

### Formato Condicional:
- Validación de asignaturas (abreviatura existe en lista)
- Validación de aulas (aula existe en catálogo)
- Alertas visuales (color rojo para errores, verde para OK)

### Estilos:
- Bordes por bloques (turnos de 2 filas)
- Colores de fondo por rango
- Encabezados estilizados
- Celdas fusionadas

---

## 7. Resumen de Capacidades

| Capacidad | Estado | Descripción |
|-----------|--------|-------------|
| Convertir JSON a Lisp | ✅ | `convertir_json_a_lisp.py` |
| Generar Python desde Lisp | ✅ | `replicar_propuesta_ods.lisp` |
| Crear Excel con formato | ✅ | `excel_gen.py` |
| Fórmulas cruzadas dinámicas | ✅ | `generador_formulas.py` |
| Hoja Aulas consolidada | ✅ | 17 grupos referenciados |
| Formato condicional | ✅ | Validación automática |
| Generación horario TV | ✅ | `generar_horario_tv_desde_lisp.py` |
| Sistema integrado (auto-detección) | ✅ | `sistema_horarios_integrado.lisp` |

---

## 8. Generación de LaTeX desde Lisp

El sistema también es capaz de generar documentos **LaTeX** directamente desde Lisp, produciendo archivos `.tex` que pueden compilarse a **PDF**. Esto se implementa en el archivo `generar_horario_tv_desde_json.tex.lisp`.

### 8.1. Archivo Principal: `generar_horario_tv_desde_json.tex.lisp`

Este archivo contiene todas las funciones necesarias para:

1. **Generar tablas de programación TV** con intervalos de tiempo a escala
2. **Generar recuadros recortables** con colores según tipo de programa
3. **Escapar caracteres especiales** para LaTeX
4. **Calcular alturas** de programas proporcionalmente a su duración

### 8.2. Sistema de Colores por Tipo de Programa

```lisp
(defparameter *tv-tipo-programa-colores*
  '(("informativo" . "blue!20")
    ("revista" . "red!20")
    ("musical" . "purple!20")
    ("infantil" . "yellow!30")
    ("cine" . "green!20")
    ("cultural" . "cyan!20")
    ("entrevista" . "orange!20")
    ("ficción" . "magenta!20")
    ("salud" . "lime!20")
    (t . "gray!10"))
  "Mapeo de tipos de programa a colores LaTeX (tonos claros).")
```

### 8.3. Función de Escape de Caracteres LaTeX

```lisp
(defun tv-tex-escape (text)
  (with-output-to-string (out)
    (loop for ch across (princ-to-string (or text ""))
          do
            (case ch
              (#\\ (princ "\\textbackslash{}" out))
              (#\& (princ "\\&" out))
              (#\% (princ "\\%" out))
              (#\$ (princ "\\$" out))
              (#\# (princ "\\#" out))
              (#\_ (princ "\\_" out))
              (#\{ (princ "\\{" out))
              (#\} (princ "\\}" out))
              (#\~ (princ "\\textasciitilde{}" out))
              (#\^ (princ "\\textasciicircum{}" out))
              (t (write-char ch out))))))
```

### 8.4. Generación de Tabla con Programas a Escala

```lisp
(defun tv-escribir-programa-escalado-tex (programa stream intervalo alto-base-ex)
  (let* ((inicio (tv-safe (getf programa :hora-inicio) ""))
         (duracion (or (getf programa :duracion) 0))
         (nombre (tv-safe (getf programa :nombre) ""))
         (inicio-min (if (> (length inicio) 0) (tv-parse-hora-a-minutos inicio) 0))
         (unidades (tv-calcular-intervalos duracion intervalo)))
    (loop for u from 0 below unidades
          for tramo-inicio = (+ inicio-min (* u intervalo))
          for tramo = (tv-rango-intervalo tramo-inicio intervalo)
          do
            (if (= u 0)
                (format stream
                        "~a & \\multirow{~a}{=}{\\parbox[c][~,1fex][c]{\\linewidth}{\\centering\\textbf{~a}}} \\\\~%"
                        (tv-tex-escape tramo)
                        unidades
                        (* alto-base-ex unidades)
                        (tv-tex-escape nombre))
                (format stream "~a &  \\\\~%" (tv-tex-escape tramo)))
            (if (< u (1- unidades))
                (format stream "\\cline{1-1}~%")
                (format stream "\\hline~%")))))
```

### 8.5. Generación de Tabla Completa por Día

```lisp
(defun tv-escribir-tabla-dia-tex (dia-config stream intervalo alto-base-ex)
  (let* ((dia (tv-dia-display (getf dia-config :dia)))
         (programas (or (getf dia-config :programas) '())))
    (format stream "\\section*{~a}~%" (tv-tex-escape dia))
    (format stream "\\begin{center}~%")
    (format stream "\\begin{adjustbox}{max totalsize={\\textwidth}{0.82\\textheight},center}~%")
    (format stream "\\begin{tabular}{|p{3.6cm}|p{11cm}|}~%")
    (format stream "\\hline~%")
    (format stream "\\textbf{Intervalo} & \\textbf{Programa} \\\\~%")
    (format stream "\\cline{1-1}~%")
    (let* ((inicio-min (tv-inicio-tabla-dia dia-config intervalo))
           (fin-min (tv-fin-tabla-dia dia-config intervalo inicio-min))
           (cantidad (max 1 (floor (- fin-min inicio-min) intervalo))))
      (loop for idx from 0 below cantidad
            for tramo-inicio = (+ inicio-min (* idx intervalo))
            for tramo = (tv-rango-intervalo tramo-inicio intervalo)
            do
              (format stream
                      "{\\centering\\rule{0pt}{~,2fex}~a} & {\\centering\\rule{0pt}{~,2fex}} \\\\~%"
                      alto-base-ex
                      (tv-tex-escape tramo)
                      alto-base-ex)
              (format stream "\\cline{1-1}~%")))
    (format stream "\\hline~%")
    (format stream "\\end{tabular}~%")
    (format stream "\\end{adjustbox}~%")
    (format stream "\\end{center}~%~%")))
```

### 8.6. Función Principal de Generación LaTeX

```lisp
(defun generar-horario-tv-tex (nombre-canal planificacion output-tex-file
                                 &key
                                   (intervalo-minutos *tv-tex-intervalo-minutos*)
                                   (alto-base-ex *tv-tex-alto-base-ex*))
  (with-open-file (stream output-tex-file
                          :direction :output
                          :if-exists :supersede)
    ;; Preámbulo LaTeX
    (format stream "\\documentclass[11pt,a4paper]{article}~%")
    (format stream "\\usepackage[utf8]{inputenc}~%")
    (format stream "\\usepackage[T1]{fontenc}~%")
    (format stream "\\usepackage[spanish]{babel}~%")
    (format stream "\\usepackage{geometry}~%")
    (format stream "\\geometry{left=1.8cm,right=1.8cm,top=1.8cm,bottom=1.8cm}~%")
    (format stream "\\usepackage{adjustbox}~%")
    (format stream "\\usepackage{multicol}~%")
    (format stream "\\usepackage{array}~%")
    (format stream "\\usepackage{xcolor}~%")
    (format stream "\\setlength{\\parindent}{0pt}~%")
    (format stream "\\begin{document}~%")
    
    ;; Título
    (format stream "\\section*{~a}~%" (tv-tex-escape nombre-canal))
    (format stream "Intervalo base: ~a minutos.\\\\~%" intervalo-minutos)
    
    ;; Generar tabla por cada día
    (loop for dia-config in planificacion
          for idx from 0
          for intervalo-dia = (tv-intervalo-dia dia-config intervalo-minutos)
          do
            (when (> idx 0) (format stream "\\clearpage~%"))
            (tv-escribir-tabla-dia-tex dia-config stream intervalo-dia alto-base-ex))
    
    (format stream "\\end{document}~%"))
  (format t "Archivo LaTeX '~a' generado.~%" output-tex-file)
  (format t "Para compilar a PDF: pdflatex ~a~%" output-tex-file)
  output-tex-file)
```

### 8.7. Generación de Recuadros Recortables (Versión Alternativa)

El sistema también puede generar una versión **recortable** donde:
1. Primera parte: tabla con celdas vacías para填写
2. Segunda parte: recuadros individuales con texto y color por tipo

```lisp
(defun tv-dibujar-rectangulo (programa stream intervalo alto-base-ex)
  (let* ((nombre (tv-safe (getf programa :nombre) ""))
         (tipo (tv-safe (getf programa :tipo-programa) ""))
         (color (tv-obtener-color-tipo tipo))
         (duracion (getf programa :duracion) 0)
         (altura-ex (max 0.8 (tv-altura-programa-ex duracion intervalo alto-base-ex))))
    (format stream
            "\\colorbox{~a}{\\fbox{\\begin{minipage}[c][~,2fex]{10.6cm}\\centering\\textbf{~a}\\end{minipage}}}~%"
            color
            altura-ex
            (tv-tex-escape nombre))
    (format stream "\\vspace{0.3cm}~%")))
```

### 8.8. Ejemplo Completo: Input → Código Lisp → Output LaTeX → PDF

**INPUT (datos del programa en Lisp):**
```lisp
(defparameter *tv-nombre-canal* "Canal Habana")

(defparameter *tv-planificacion-semanal*
  '(
    (:dia "lunes" :programas
     ((:nombre "HABANA NOTICIARIO" :duracion 30 :hora-inicio "17:30" :hora-final "18:00" :tipo-programa "informativo")
      (:nombre "POWER RANGERS" :duracion 30 :hora-inicio "18:00" :hora-final "18:30" :tipo-programa "infantil")))
    (:dia "martes" :programas ...)))
```

**PROCESO (ejecutar en SBCL):**
```bash
sbcl --eval "(load \"generar_horario_tv_desde_json.tex.lisp\")" \
     --eval "(generar-horario-tv-desde-json :output-tex-file \"horario_tv.tex\")" \
     --quit

# Compilar a PDF
pdflatex horario_tv.tex
```

**OUTPUT (horario_tv.tex - fragmento):**
```latex
\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[spanish]{babel}
\usepackage{geometry}
\geometry{left=1.8cm,right=1.8cm,top=1.8cm,bottom=1.8cm}
\usepackage{adjustbox}
\usepackage{multicol}
\usepackage{array}
\usepackage{xcolor}
\begin{document}
\section*{Canal Habana}
Intervalo base: 15 minutos.\\
La columna Intervalo es fija por unidad; Programa se dibuja a escala.\par

\section*{Lunes}
\begin{center}
\begin{adjustbox}{max totalsize={\textwidth}{0.82\textheight},center}
\begin{tabular}{|p{3.6cm}|p{11cm}|}
\hline
\textbf{Intervalo} & \textbf{Programa} \\
\hline
{\centering\rule{0pt}{5.60ex}16:00--16:15} & {\centering\rule{0pt}{5.60ex}} \\
...
{\centering\rule{0pt}{5.60ex}17:30--17:45} & {\centering\rule{0pt}{5.60ex}} \\
\hline
\end{tabular}
\end{adjustbox}
\end{center}
\clearpage
...
\end{document}
```

**OUTPUT (PDF generado):**
```
┌─────────────────────────────────────────────────────────────────┐
│                      CANAL HABANA                               │
│Intervalo base: 15 minutos.                                     │
│                                                                  │
│────────────────────────── LUNES ─────────────────────────────── │
│ Intervalo    │ Programa                                         │
│──────────────┼─────────────────────────────────────────────────│
│ 16:00-16:15 │                                                  │
│ 16:15-16:30 │                                                  │
│ ...         │                                                  │
│ 17:30-17:45 │                                                  │
│──────────────┼─────────────────────────────────────────────────│
```

### 8.9. Comparación: Las Tres Versiones de Output

| Versión | Descripción | Uso |
|---------|-------------|-----|
| **Excel** (.xlsx) | Hojas de cálculo con fórmulas dinámicas | Para edición y cálculo automático |
| **LaTeX básico** (.tex) | Tabla con programas a escala | Para imprimir guía de programación |
| **LaTeX recortable** (.tex) | Tabla vacía + recuadros coloridos | Para cortar y pegar en producción |

### 8.10. Comandos de Ejecución para LaTeX

```bash
# Generar LaTeX desde Lisp (tabla básica)
sbcl --eval "(load \"generar_horario_tv_desde_json.tex.lisp\")" \
     --eval "(generar-horario-tv-desde-json :output-tex-file \"horario_tv.tex\" :intervalo-minutos 15)" \
     --quit

# Generar versión recortable
sbcl --eval "(load \"generar_horario_tv_desde_json.tex.lisp\")" \
     --eval "(generar-horario-tv-desde-json :generar-recortable t)" \
     --quit

# Compilar a PDF
pdflatex horario_tv.tex
pdflatex horario_recortable.tex
```

---

## 9. Resumen de Capacidades del Sistema Completo

### 9.1. Pipeline de Generación

```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│   JSON      │───►│    LISP     │───►│   Python    │───►│   Excel     │
│ (datos)     │    │ (generador) │    │ (motor)     │    │ (.xlsx)     │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
                        │
                        ▼
                   ┌─────────────┐    ┌─────────────┐
                   │   LaTeX     │───►│    PDF      │
                   │ (.tex)      │    │             │
                   └─────────────┘    └─────────────┘
```

### 9.2. Capacidades por Tipo de Output

| Capacidad | Excel | LaTeX/PDF |
|-----------|-------|-----------|
| Generación desde JSON | ✅ | ✅ |
| Metaprogramación con Lisp | ✅ | ✅ |
| 17+ grupos académicos | ✅ | ❌ |
| Fórmulas dinámicas cruzadas | ✅ | ❌ |
| Formato condicional | ✅ | ❌ |
| Tabla con programas a escala | ✅ | ✅ |
| Recuadros recortables | ❌ | ✅ |
| Colores por tipo de programa | ❌ | ✅ |
| Compilación a PDF | ❌ | ✅ |

### 9.3. Archivos Principales del Sistema

| Archivo | Rol |
|---------|-----|
| `convertir_json_a_lisp.py` | Convierte JSON a variables Lisp |
| `codigo-tesis.lisp` | Framework de clases (tabla, hoja, libro) |
| `variables_horario.lisp` | Datos de horarios de 17 grupos |
| `replicar_propuesta_ods.lisp` | Orquestador: genera Python desde Lisp |
| `sistema_horarios_integrado.lisp` | Sistema integrado con auto-detección |
| `generar_horario_tv_desde_json.tex.lisp` | Generador de LaTeX y Python para TV |
| `excel_gen.py` / `hoja_con_formulas.py` | Motor de generación Excel |
| `generador_formulas.py` | Generador de fórmulas cruzadas |

---

*Informe generado para la tesis de graduación*
*Sistema de Generación de Horarios - MATCOM*