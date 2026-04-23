# Tutorial: usar `hoja_con_formulas.py` para horarios

Este tutorial te deja un flujo listo en 2 fases:

- **Fase 1**: solo pasas nombres de grupos y se crean hojas de grupos vacías + hoja `Aulas` con fórmulas cruzadas.
- **Fase 2**: además pasas asignaturas por grupo y se completa la tabla de asignaturas de cada hoja.

El archivo base del tutorial es: `tutorial_hoja_con_formulas.py`.

---

## 1) Qué se genera

Al ejecutar el script se crean dos Excel:

- `tutorial_fase1_grupos_vacios.xlsx`
- `tutorial_fase2_con_asignaturas.xlsx`

Cada hoja de grupo incluye:

- tabla de horario (vacía)
- tabla de asignaturas
- tabla lateral de aulas
- fórmulas de totales y ocupación
- formato condicional
- separador entre **Turno 3** y **Turno 4**
- bordes en la fila separadora

La hoja `Aulas` incluye:

- bloques por día
- fórmulas cruzadas que leen todas las hojas de grupos

---

## 2) Ejecutar

Desde la raíz del proyecto:

```bash
python3 tutorial_hoja_con_formulas.py
```

---

## 3) Fase 1: solo nombres de grupos

En `tutorial_hoja_con_formulas.py`, modifica esta lista:

```python
grupos = ["D111", "D211", "C111"]
```

Y llama:

```python
config_fase_1 = build_config(grupos)
generar_excel_personalizado(config_fase_1, "tutorial_fase1_grupos_vacios.xlsx")
```

Sí, en esta fase basta con los nombres de grupos.

---

## 4) Fase 2: grupos + asignaturas

Pasa un diccionario `asignaturas_por_grupo`:

```python
asignaturas_por_grupo = {
    "D111": [
        ("AL", "Álgebra Lineal", 3),
        ("L", "Lógica", 2),
    ],
    "D211": [
        ("MA", "Matemática y Aplicaciones", 2),
    ],
}

config_fase_2 = build_config(grupos, asignaturas_por_grupo)
generar_excel_personalizado(config_fase_2, "tutorial_fase2_con_asignaturas.xlsx")
```

También puedes usar `dict` en cada asignatura:

```python
{"abrev": "AL", "asignatura": "Álgebra Lineal", "frec": 3}
```

---

## 5) API útil del tutorial

Funciones principales en `tutorial_hoja_con_formulas.py`:

- `build_config(grupos, asignaturas_por_grupo=None)`
- `build_group_sheet_config(group, subjects=None, aulas_catalogo=None)`
- `build_aulas_sheet_config(groups)`
- `build_aulas_fernando_formulas(groups, row_step=3, turnos=6)`

Con eso puedes integrar el generador a otros scripts sin tocar `hoja_con_formulas.py`.

---

## 6) Parámetros globales que puedes ajustar

Al inicio del script:

- `TURNOS = 6`
- `HORARIO_ROW_STEP = 3`
- `AULAS_CATALOGO = ["Aula 1", ..., "Lab"]`
- colores y estilos (`COLOR_*`)

---

## 7) Validaciones rápidas recomendadas

- Verifica que exista separador vacío con borde en fila 13 (`B13:G13`) de cada hoja de grupo.
- Verifica que `Aulas` tenga fórmulas en bloques `C4:L9`, `C12:L17`, `C20:L25`, `C28:L33`, `C36:L41`.
- Abre Excel/LibreOffice y recalcula fórmulas si no se actualizan de inmediato.

---

## 8) Siguiente paso sugerido

Si quieres conectarlo a tus datos reales desde Lisp/JSON, usa `build_config(...)` como capa intermedia y solo alimenta:

- `grupos`
- `asignaturas_por_grupo`

Todo lo demás (rangos, fórmulas, condicionales, bordes y hoja `Aulas`) se arma automáticamente.
