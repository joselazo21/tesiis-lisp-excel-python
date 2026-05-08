# Informe de Migración: Nuevo Sistema de Generación Excel

## Resumen Ejecutivo

Se ha migrado el sistema de generación de archivos Excel desde un enfoque basado en funciones que generan código Python directamente, hacia una nueva arquitectura basada en clases que describen estructuras tabulares de datos. Cada nodo del AST genera código Python que utiliza la biblioteca `openpyxl` para crear archivos Excel.

## Objetivo Anterior vs Nuevo

### Enfoque Anterior
- Funciones Lisp que generan código Python inline
- Acoplamiento directo entre la lógica de generación y el output
- Difícil de extender y mantener

### Nuevo Enfoque
- Clases que modelan la estructura de datos tabulares
- Métodos `generate-code` separados por tipo de nodo
- clean separation between model and code generation

## Clases Definidas

### Clases Core

| Clase | Descripción |
|-------|------------|
| `celda` | Celda individual con valor, fórmula, estilo y validación |
| `fila` | Fila con índice, nombre y celdas |
| `columna` | Columna con ancho, tipo de dato y restricciones |
| `style` | Estilo visual: fuente, colores, alineación, bordes |
| `border-style` | Configuración de bordes |
| `formato-condicional` | Reglas de formato condicional |
| `validacion` | Validación de datos de celda |
| `restriccion` | Restricciones de columna (requerido, patrón, rango) |

### Clases de Estructura

| Clase | Descripción |
|-------|------------|
| `relacion` | Relaciones entre tablas (clave foránea) |
| `tabla-modelo` | Estructura tabular con contenido, estilos, formatos |
| `hoja-modelo` | Hoja de cálculo con tablas, fórmulas y relaciones |
| `libro-modelo` | Libro completo con hojas y metadata |

## Métodos generate-code

Cada clase tiene un método `generate-code` que genera código Python apropiado:

```lisp
(defmethod generate-code ((node clase-tabla-modelo)
                        (lang output-python-openpyxl)
                        (stream t))
  ;; Genera configuración de tabla en Python
```

### Métodos Implementados

1. **clase-celda** → Genera valor y fórmula de celda
2. **clase-style** → Genera diccionario de estilos
3. **clase-border-style** → Genera configuración de bordes
4. **clase-validacion** → Genera reglas de validación
5. **clase-restriccion** → Genera restricciones de columna
6. **clase-formato-condicional** → Genera reglas de formato condicional
7. **clase-tabla-modelo** → Genera datos de tabla y configuraciones
8. **clase-hoja-modelo** → Genera código de hoja completa
9. **clase-libro-modelo** → Genera script Python completo

##Salida Generada

### Estructura del Script Python

```python
#!/usr/bin/env python3
"""Generador Excel:HorarioAulas
Generado desde modelo-excel.clisp"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import Rule

# ===== DATOS =====
Grupo11_tabla = [...]

# ===== FUNCIONES =====
def apply_style(ws, cell_range, style_dict):
    ...

def add_conditional_format(ws, cell_range, formula, rule_type, color):
    ...

# ===== GENERADOR =====
def generar_excel_HorarioAulas():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "HorarioAulas"
    ...

def main():
    wb = generar_excel_HorarioAulas()
    wb.save("HorarioAulas.xlsx")
    print(f'Archivo: HorarioAulas.xlsx')

if __name__ == '__main__':
    main()
```

## Funciones Helper

### Generación de Referencias Excel

```lisp
(python-cell-ref 1 1)  ; → "A1"
(python-range-ref 1 1 3 5)  ; → "A1:C5"
(numero-a-letra-columna 1)  ; → "A"
```

### Escritura de Valores Python

```lisp
(write-python-value '(1 2 3) stream)  ; → [1, 2, 3]
(write-python-value "hola" stream)  ; → "hola"
(write-python-value nil stream)  ; → None
```

### Generación de Fórmulas

```lisp
(generate-excel-formula "CONCAT" "A1" "B1")  ; → "=CONCAT(A1, B1)"
(generate-excel-formula "IF" "A1>0" "positivo" "negativo")  ; → "=IF(A1>0, positivo, negativo)"
```

## Macros de Creación

```lisp
;; Crear tabla simple
(crear-tabla-simple "asignaturas" '(("Algebra" 1) ("Logica" 1)))

;; Crear hoja con tablas
(crear-hoja-simple "Grupo11" tabla1 tabla2)

;; Crear libro completo
(crear-libro-excel "HorarioAulas" hoja1 hoja2)
```

## Uso

### Cargar el sistema
```lisp
(load "modelo-excel.clisp")
```

### Generar ejemplo básico
```lisp
(ejemplo-generar-libro)
```

### Generar y guardar archivo
```lisp
(generar-archivo-excel (ejemplo-generar-libro) "test_generado.py")
```

### Crear personalizado
```lisp
(let ((tabla (make-instance 'clase-tabla-modelo
                          :id "horario-lunes"
                          :contenido '(("Turno 1" "Aula 101")
                                     ("Turno 2" "Aula 102")))))
  (generate-code tabla python-openpyxl t))
```

## Beneficios de la Nueva Arquitectura

1. **Extensibilidad**: Agregar nuevas clases sin modificar código existente
2. **Mantenibilidad**: Cada clase responsables de su propia generación
3. **Testabilidad**: Cada nodo puede probarse independientemente
4. **Reusabilidad**: Las mismas clases pueden generar diferentes formatos
5. **Type Safety**: Estructuras definidas claramente

## Siguientes Pasos

1. Migrar las funciones existentes de `replicar_propuesta_ods.lisp` para usar estas clases
2. Agregar soporte para múltiples hojas en el libro
3. Implementar fórmulas dinámicas entre celdas
4. Agregar validación de datos más completa
5. Soporte para gráficos y tablas pivote

## Archivos

- **modelo-excel.clisp** (~680 líneas): Nuevo sistema de clases y generación
- **codigo-tesis.lisp** (dependencia): Clases base y utilities

---

*生成: Mayo 2026*
*Autor: opencode*