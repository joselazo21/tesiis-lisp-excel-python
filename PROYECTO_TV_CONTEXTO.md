# Proyecto: Sistema de Generación de Horarios TV

## Contexto General

Sistema para generar archivos Excel de programación TV semanal desde datos JSON. Utiliza **Common Lisp** para describir estructuras de datos y **Python/openpyxl** para generar archivos Excel.

## Arquitectura

```
[Datos JSON] → [Lisp (model)] → [Python (openpyxl)] → [Excel .xlsx]
```

### Stack Tecnológico
- **Lisp**: clisp o SBCL
- **Librería Excel**: openpyxl (Python)
- **Idioma**: Español

## Archivos Clave

### Core del Sistema
| Archivo | Propósito |
|--------|-----------|
| `codigo-tesis.lisp` | Macro `defclass*` y utilidades base |
| `modelo-excel.lisp` | Clases: `xl-table`, `xl-sheet`, `xl-workbook` |
| `generar_tv_modelo.lisp` | Funciones TV: `crear-tabla-tv-dia`, etc. |
| `variables_horario_tv.lisp` | Datos de programación (JSON transformado) |

### Scripts de Generación
| Archivo | Propósito |
|--------|----------|
| `generar_tv.sh` | Shell script principal (usa SBCL) |
| `test_tv_simple.sh` | Prueba simplificada |

### Outputs
| Archivo | Propósito |
|--------|----------|
| `horario_tv.py` | Python generado |
| `horario_tv.xlsx` | Excel final |

## Clases del Modelo (modelo-excel.lisp)

```lisp
(defclass* xl-table () (id cols rows contenido))
(defclass* xl-sheet () (name tables))
(defclass* xl-workbook () (name sheets))
```

### Cómo Usar

```lisp
;; Crear una tabla
(xl-table :id "tv_lunes" 
          :contenido '(("16:00" "16:05" 5 "Noticias" "informativo")))

;; Crear una hoja con tablas
(xl-sheet :name "Lunes"
          :tables (list (xl-table ...)))

;; Crear workbook
(xl-workbook :name "Canal Habana"
             :sheets (list (xl-sheet ...)))

;; Generar Python
(xl-generate workbook "salida.py")
```

## Regla Crítica: Argumentos con Keyword

**IMPORTANTE**: Los constructores creados por `defclass*` usan `&key`. **Debes usar keywords**:

```lisp
;; ✓ Correcto
(xl-table :id "tabla1" :contenido datos)

;; ✗ Incorrecto (argumentos posicionales)
(xl-table "tabla1" nil nil datos)
```

Si usas argumentos posicionales, los datos no se asignarán correctamente.

## Flujo de Generación

1. **Cargar**: `(load "codigo-tesis.lisp")` → `modelo-excel.lisp` → `generar_tv_modelo.lisp`
2. **Datos**: Cargar `variables_horario_tv.lisp` (contiene `*tv-planificacion-semanal*`)
3. **Generar**: `(generar-horario-tv :planificacion *tv-planificacion-semanal* :nombre-canal "Canal Habana" :output-file "horario_tv.py")`
4. **Ejecutar**: `python3 horario_tv.py` → genera `horario_tv.xlsx`

## Errores Comunes

| Error | Causa | Solución |
|-------|------|---------|
| Python generado vacío | Usar argumentos posicionales en lugar de keywords | Usar `:id`, `:contenido`, `:name`, etc. |
| SBCL keyword errors | `defclass*` crea constructores `&key` | Usar clisp en lugar de SBCL, o usar keyword args |
| "unhandled condition" | SBCL no maneja keyword args bien | Usar clisp: `clisp -q script.lisp` |

## Estructura de Datos TV

```lisp
(defparameter *tv-planificacion-semanal*
  '((:dia "lunes"
      :programas ((:nombre " Programa" 
                   :duracion 30
                   :hora-inicio "16:00"
                   :hora-final "16:30"
                   :tipo-programa "informativo"
                   :tipo-publico "adulto")
                  ...))))
```

## Para Continuar una Sesión

1. **Cargar el modelo**:
   ```lisp
   (load "codigo-tesis.lisp")
   (load "modelo-excel.lisp")
   (load "generar_tv_modelo.lisp")
   (load "variables_horario_tv.lisp")
   ```

2. **Generar**:
   ```lisp
   (generar-horario-tv :planificacion *tv-planificacion-semanal*
                     :nombre-canal *tv-nombre-canal*
                     :output-file "horario_tv.py")
   ```

3. **Ejecutar Python**:
   ```bash
   python3 horario_tv.py
   ```

## Comandos Útiles

```bash
# Usar clisp (recomendado para evitar errores de keyword)
clisp -q test_tv.lisp

# Generar con sh
bash generar_tv.sh
```

## Estado Actual

- ✓ Migration completada: funciones → modelo AST
- ✓ Keyword args funcionan correctamente
- ✓ 7 días generados con datos completos
- ✓ Python genera Excel correctamente