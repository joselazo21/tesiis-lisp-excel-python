# Jerarquía de Clases xl-

```
╔══════════════════════════════════════════════════════════════════════════════════════╗
║                          JERARQUÍA DE CLASES xl-                              ║
╠══════════════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║                              xl-out (marker)                                  ║
║                              ╔═══════════╗                                    ║
║                              ║ xl-out    ║                                    ║
║                              ╚═══════════╝                                    ║
║                                   ▲                                         ║
║         ╔════════════════════════════╬════════════════════════════╗            ║
║         │                            │                        │              ║
║         │ xl-workbook               │ xl-sheet              xl-table        ║
║   ┌─────┴─────┐            ┌──────┴──────┐       ┌─────┴─────┐         ║
║   │ name     │            │ name       │       │ contenido │         ║
║   │ sheets   │            │ tables     │       │ headers   │         ║
║   └─────────┘            │ formulas   │       │ column-   │         ║
║         │               │ f.do-formas│        │ widths   │         ║
║         │               │ table-    │        └─────────┘         ║
║         │               │ borders   │                         ║
║         │               │ border-   │                         ║
║         │               │ color/style                         ║
║         │               │ table-    │                         ║
║         │               │ ranges   │                         ║
║         │               │ range-   │                         ║
║         │               │ styles   │                         ║
║         │               │ merge-   │                         ║
║         │               │ ranges  │                         ║
║         │               │ cond.   │                         ║
║         │               │ format- │                         ║
║         │               │ rules   │                         ║
║         │               │ cell-   │                         ║
║         │               │ size    │                         ║
║         │               │ header- │                         ║
║         │               │ style   │                         ║
║         │               └─────────┘                         ║
║         │                                                  ║
║         ▼                                                  ║
║   ═══════════════════════════════════════════════════════���══  ║
║                    CLASES DE ESTILO                          ║
║   ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  ║
║   │ xl-style     │  │ xl-range-    │  │ xl-header-style  │  ║
║   │              │  │ style       │  │                  │  ║
║   │ bold        │  │ range       │  │ bold             │  ║
║   │ color       │  │ style       │  │ color            │  ║
║   │ bg-color    │  └──────────────┘  │ bg-color         │  ║
║   │ align       │                    │ align            │  ║
║   └──────────────┘                    └──────────────────┘  ║
║                                                                              ║
║   ═══════════════════════════════════════════════════════════  ║
║                    CLASES DE FÓRMULA                          ║
║   ┌──────────────────┐  ┌─────────────────────┐            ║
║   │ xl-formula       │  │ xl-fernando-formula  │            ║
║   │                 │  │                     │            ║
║   │ row             │  │ cell (ref: "C4")    │            ║
║   │ col             │  │ formula (string)     │            ║
║   │ value (=SUM...) │  │                     │            ║
║   └──────────────────┘  └─────────────────────┘            ║
║                                                                              ║
║   ═══════════════════════════════════════════════════════════  ║
║             OTRAS CLASES                                     ║
║   ┌──────────────────────┐  ┌─────────────────┐              ║
║   │ xl-conditional-rule │  │ xl-merge-range  │              ║
║   │                    │  │                 │              ║
║   │ tipo               │  │ range ("B4:B5") │              ║
║   │ rango              │  └─────────────────┘              ║
║   │ formula                                                      ║
║   │ color                                                       ║
║   └──────────────────────┘                                      ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════════════╝
```

---

## Leyenda

| Símbolo | Significado |
|---------|-------------|
| `┌─────────┐` | `defclass*` (constructor usando `:keyword` args) |
| `│         │` | slots (atributos de cada clase) |
| `└─────────┘` | Relación de herencia |
| `▲` | Herencia hacia arriba |

---

## Ejemplo de Uso

```lisp
;; 1. Crear tabla con datos y headers
(xl-table :contenido '(("16:00" "16:30" 30 "Noticias"))
           :headers '("Inicio" "Fin" "Duración" "Programa"))

;; 2. Crear hoja con estilo
(xl-sheet :name "Lunes"
          :tables (list tabla)
          :table-borders t
          :border-color "B7B7B7"
          :header-style (xl-header-style :bold t :bg-color "4A90E2"))

;; 3. Crear workbook
(xl-workbook :name "Test.xlsx"
             :sheets (list hoja))

;; 4. Generar Python
(xl-generate workbook "salida.py")
```

---

## Flujo de Generación

```
┌─────────────────┐
│  xl-workbook    │  ← Objeto Lisp con :name, :sheets
└────────┬────────┘
         │ generate-code
         ▼
┌─────────────────────────────┐
│  Python config dict        │  ← Diccionario Python
│  {                       │
│    "sheets": [...]       │
│  }                       │
└────────┬────────────────┘
         │
         ▼
┌─────────────────────────────────────┐
│  hoja_con_formulas.py                │  ← Framework Python
│  generar_excel_personalizado()       │
│         │                          │
│         ▼ (openpyxl)              │
└────────┬──────────────────────────┘
         │
         ▼
┌─────────────────┐
│  Test.xlsx     │  ← Archivo Excel
└─────────────────┘
```

---

## Slots por Clase

### xl-table
- `id` - identificador
- `contenido` - datos (lista de listas)
- `headers` - nombres de columnas
- `column-widths` - anchos de columnas

### xl-sheet
- `name` - título de la hoja
- `tables` - lista de tablas
- `formulas` - fórmulas simples
- `fernando-formulas` - fórmulas complejas cross-sheet
- `table-borders` - bordes (t/nil)
- `border-color` - color de borde
- `border-style` - estilo de borde
- `table-ranges` - rangos para bordes
- `range-styles` - estilos por rango
- `merge-ranges` - celdas fusionadas
- `conditional-format-rules` - formato condicional
- `column-widths` - anchos de columnas
- `cell-size` - factor de escala
- `header-style` - estilo de headers

### xl-workbook
- `name` - nombre del archivo
- `sheets` - lista de hojas