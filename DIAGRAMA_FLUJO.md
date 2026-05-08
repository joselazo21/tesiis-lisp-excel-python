# Diagrama de Flujo: Sistema de Generación de Horarios TV

## Flujo Principal

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                          INICIO                                            │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  1. DATOS DE ENTRADA                                                       │
│     variables_horario_tv.lisp                                               │
│                                                                              │
│     *tv-planificacion-semanal*  ←  JSON transformado                        │
│     *tv-nombre-canal*          ←  "Canal Habana"                           │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  2. CARGAR LIBRERÍAS EN LISP                                                │
│                                                                              │
│     (load "codigo-tesis.lisp")     →  Macro defclass*                       │
│     (load "modelo-excel.lisp")     →  Clases xl-*                            │
│     (load "generar_tv_modelo.lisp")→  Funciones TV                          │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  3. CONSTRUCCIÓN DEL MODELO (AST)                                            │
│                                                                              │
│  ┌─────────────────────────────────────────┐                                │
│  │ generar-horario-tv                       │                                │
│  │   │                                    │                                │
│  │   ▼                                    │                                │
│  │ crear-libro-tv                          │                                │
│  │   │  :nombre-canal                     │                                │
│  │   │  :planificacion                    │                                │
│  │   ���                                    │                                │
│  │ xl-workbook                            │◄── ┐                             │
│  │   └── sheets: list of                 │    │                             │
│  │            │                         │    │                             │
│  │            ▼                         │    │                             │
│  │       crear-hoja-tv-dia  (x7)         │    │                             │
│  │            │                         │    │                             │
│  │            ▼                         │    │                             │
│  │       xl-sheet                       │    │                             │
│  │            └── tables                │    │                             │
│  │                 │                  │    │                             │
│  │                 ▼                  │    │                             │
│  │            crear-tabla-tv-dia        │    │                             │
│  │                 │                   │    │                             │
│  │                 ▼                   │    │                             │
│  │            xl-table                 │    │                             │
│  │                 │                   │    │                             │
│  │                 ▼                   │    │                             │
│  │            programa-tv-a-fila       │    │                             │
│  │                 │                   │    │                             │
│  │                 ▼                   │    │                             │
│  │            contenido: list of rows   │    │                             │
│  └───────────────────────────────────────┘    │                             │
│                                              │                             │
└───────────────────────────────────────────────│─────────────────────────────┘
                                               │
                                               ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  4. GENERACIÓN DE CÓDIGO PYTHON                                            │
│                                                                              │
│     xl-generate (workbook, "horario_tv.py")                                  │
│                                                                              │
│     generate-code para cada clase:                                            │
│                                                                              │
│     ┌─────────────────────────────────────────┐                             │
│     │ xl-workbook.generate-code               │                             │
│     │   ├──Genera: #!/usr/bin/env python3      │                             │
│     │   ├──Import: openpyxl                   │                             │
│     │   ├──Data vars (xl_*_data = [...])      │                             │
│     │   ├──generar() function               │                             │
│     │   │   ├──wb = Workbook()             │                             │
│     │   │   ├──ws = wb.active               │                             │
│     │   │   └──data loops                   │                             │
│     │   └──main() function                 │                             │
│     └─────────────────────────────────────────┘                             │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  5. OUTPUT: horario_tv.py                                                  │
│                                                                              │
│     #!/usr/bin/env python3                                                  │
│     import openpyxl                                                         │
│                                                                          │
│     # DATA                                                                │
│     tv_lunes_data = [["16:00","16:05",5,"Noticias",...], ...]                │
│     tv_martes_data = [...], ...                                              │
│                                                                          │
│     def generar():                                                        │
│         wb = openpyxl.Workbook()                                           │
│         ws = wb.active                                                    │
│         ws.title = "Canal Habana"                                         │
│         for r,d in enumerate(tv_lunes_data,1):                          │
│             for c,v in enumerate(d,1):                                  │
│                 ws.cell(r,c,v)                                           │
│         ...                                                               │
│         return wb                                                        │
│                                                                          │
│     def main():                                                           │
│         wb = generar()                                                   │
│         wb.save('horario_tv.xlsx'); print('OK')                          │
│                                                                          │
│     if __name__=='__main__': main()                                        │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  6. EJECUCIÓN PYTHON                                                       │
│                                                                              │
│     python3 horario_tv.py                                                 │
│                                                                              │
│     ├── openpyxl.Workbook()  →  Creates in memory                           │
│     ├── ws.cell(r,c,v)        →  Fills cells                                 │
│     └── wb.save()          →  Writes horario_tv.xlsx                       │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  7. OUTPUT FINAL: horario_tv.xlsx                                         │
│                                                                              │
│     ┌──────────┬──────────┬──────────┬──────────┬──────────┬──────────┐   │
│     │   A      │    B    │    C    │    D     │    E     │    F     │   │
│  ───┼──────────┼──────────┼──────────┼─────────���┼──────────┼──────────┤   │
│  1  │ Hora Ini │ Hora Fin│Duración │ Programa │  Tipo    │ Público │   │
│  ───┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┤   │
│  2  │ 16:00   │ 16:05   │    5    │Noticias  │inform.  │ adulto  │   │
│  3  │ 16:05   │ 16:10   │    5    │Coords.  │inform.  │ adulto  │   │
│  4  │ ...     │ ...     │   ...   │ ...     │  ...    │  ...    │   │
│     └──────────┴──────────┴──────────┴──────────┴──────────┴──────────┘   │
│                                                                              │
│                     FIN                                                     │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Diagrama de Componentes (Arquitectura)

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        APLICACIÓN LISP                                   │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │                    modelo-excel.lisp                            │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐              │   │
│  │  │ xl-table    │  │ xl-sheet    │  │xl-workbook  │              │   │
│  │  │  :id        │  │  :name     │  │  :name      │              │   │
│  │  │  :contenido │  │  :tables   │  │  :sheets    │              │   │
│  │  └─────────────┘  └─────────────┘  └─────────────┘              │   │
│  │         │                │                │                      │   │
│  │         └────────┬───────┴───────┬───────┘                      │   │
│  │                  │generate-code│                              │   │
│  └──────────────────│────────────│──────────────────────────────┘   │
│                     │            │                                      │
└─────────────────────│────────────│──────────────────────────────────────┘
                      ▼            ▼
         ┌────────────────────────────┐
         │    PYTHON GENERADO        │
         │    (horario_tv.py)       │
         │                         │
         │  import openpyxl         │
         │  tv_lunes_data = [...]   │
         │  def generar(): ...     │
         │  def main(): ...      │
         └───────────┬────────────┘
                     │
                     ▼
         ┌��─��─────────────────────────┐
         │    openpyxl               │
         │    (runtime)              │
         │                         │
         │  Workbook → Worksheet   │
         │  cells → save .xlsx       │
         └───────────┬────────────┘
                     │
                     ▼
         ┌────────────────────────────┐
         │    horario_tv.xlsx       │
         │    (Archivo Excel)       │
         └────────────────────────┘
```

---

## Ciclo de Vida de un Dato

```
1. PROGRAMA Lisp (en variables_horario_tv.lisp)
   ┌──────────────────────────────────────┐
   │ (:nombre "EL TIEMPO Y LA MEMORIA"     │
   │  :duracion 5                        │
   │  :hora-inicio "16:00"               │
   │  :hora-final "16:05")                │
   └──────────────────────────────────────┘
                    │
                    ▼
2. programa-tv-a-fila (transforma a lista)
   ┌──────────────────────────────────────┐
   │ ("16:00" "16:05" 5                  │
   │  "EL TIEMPO Y LA MEMORIA"            │
   │  "informativo" "adulto")             │
   └──────────────────────────────────────┘
                    │
                    ▼
3. xl-table :contenido (almacena lista)
   ┌──────────────────────────────────────┐
   │ contenido: (("16:00" "16:05" 5     │
   │              "EL TIEMPO..." ...)...) │
   └──────────────────────────────────────┘
                    │
                    ▼
4. generate-code (genera Python)
   ┌──────────────────────────────────────┐
   │ tv_lunes_data = [["16:00","16:05",5,   │
   │   "EL TIEMPO Y LA MEMORIA",           │
   │   "informativo","adulto"], ...]      │
   └──────────────────────────────────────┘
                    │
                    ▼
5. Python runtime (ws.cell)
   ┌──────────────────────────────────────┐
   │ ws.cell(2, 1, "16:00")               │
   │ ws.cell(2, 2, "16:05")               │
   │ ws.cell(2, 3, 5)                     │
   │ ws.cell(2, 4, "EL TIEMPO...")        │
   └──────────────────────────────────────┘
                    │
                    ▼
6. Excel cell (A2)
   ┌──────────────────────────────────────┐
   │  16:00                               │
   └──────────────────────────────────────┘
```