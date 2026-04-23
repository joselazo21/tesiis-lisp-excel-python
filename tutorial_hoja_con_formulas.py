"""Tutorial práctico para generar horarios con hoja_con_formulas.py.

Fase 1:
  - Solo pasas nombres de grupos.
  - Se generan hojas de grupos con tablas vacías, fórmulas y formato condicional.
  - Se genera hoja Aulas con fórmulas cruzadas a cada grupo.

Fase 2:
  - Además pasas asignaturas por grupo.
  - Se rellenan las tablas de asignaturas y se activan fórmulas por fila.
"""

from __future__ import annotations

from typing import Iterable

from hoja_con_formulas import generar_excel_personalizado


# =========================
# Configuración base
# =========================

DAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
AULAS_CATALOGO = [f"Aula {i}" for i in range(1, 10)] + ["Lab"]
TURNOS = 6
HORARIO_ROW_STEP = 3

COLOR_VERDE_ASIGNATURAS = "A9D18E"
COLOR_ROJO_TURNOS = "F4CCCC"
COLOR_ROJO_AULAS = "E6B8AF"
COLOR_BORDE = "4F81BD"


def col_letter(col_num: int) -> str:
    """Convierte número de columna (1-based) a letra Excel."""
    value = ""
    n = col_num
    while n > 0:
        n, rem = divmod(n - 1, 26)
        value = chr(65 + rem) + value
    return value


def build_row_names(turnos: int = TURNOS, row_step: int = HORARIO_ROW_STEP) -> list[str]:
    """Crea los nombres de filas del horario con separador entre Turno 3 y 4."""
    rows: list[str] = []
    for turno in range(1, turnos + 1):
        rows.append(f"Turno {turno}")
        rows.extend([""] * (row_step - 1))
        if turno == 3:
            rows.append("")
    return rows


def normalize_subject(item: object) -> list[object]:
    """Normaliza una asignatura a formato [abrev, nombre, frec, faltan, asignadas]."""
    if isinstance(item, dict):
        abrev = item.get("abrev", "")
        nombre = item.get("asignatura", item.get("nombre", ""))
        frec = item.get("frec", item.get("frecuencia", ""))
    elif isinstance(item, (list, tuple)) and len(item) >= 3:
        abrev, nombre, frec = item[0], item[1], item[2]
    else:
        raise ValueError(
            "Formato de asignatura inválido. Usa tuple/list (abrev, nombre, frec) o dict con abrev/asignatura/frec."
        )
    return [abrev, nombre, frec, "", ""]


def build_dynamic_merge_ranges(turnos: int = TURNOS, row_step: int = HORARIO_ROW_STEP) -> list[str]:
    """Rangos merge para columna B, respetando separador entre Turno 3 y 4."""
    ranges: list[str] = []
    for turno in range(1, turnos + 1):
        offset = 1 if turno >= 4 else 0
        start_row = 4 + (turno - 1) * row_step + offset
        end_row = start_row + (row_step - 1)
        ranges.append(f"B{start_row}:B{end_row}")
    return ranges


def build_group_sheet_config(
    group: str,
    subjects: Iterable[object] | None = None,
    aulas_catalogo: list[str] | None = None,
) -> dict:
    """Construye la configuración de una hoja de grupo."""
    if subjects is None:
        subjects = []
    if aulas_catalogo is None:
        aulas_catalogo = AULAS_CATALOGO

    row_names = build_row_names()
    total_horario_rows = len(row_names)
    horario_end_row = 3 + total_horario_rows  # 22

    expected_turno4_index = 3 * HORARIO_ROW_STEP
    actual_turno4_index = row_names.index("Turno 4")
    pre_end_row = 3 + expected_turno4_index   # 12
    post_start_row = 4 + actual_turno4_index  # 14
    separator_row = 4 + expected_turno4_index  # 13

    dynamic_turnos_ranges = [f"B4:B{pre_end_row}", f"B{post_start_row}:B{horario_end_row}"]
    dynamic_horario_ranges = [f"C4:G{pre_end_row}", f"C{post_start_row}:G{horario_end_row}"]
    dynamic_separator_range = f"B{separator_row}:G{separator_row}"

    normalized_subjects = [normalize_subject(s) for s in subjects]
    asig_height = max(1, len(normalized_subjects))
    asig_end_row = 3 + asig_height
    dynamic_asig_range = f"I4:M{asig_end_row}"
    asig_abrev_range = f"$I4:I{asig_end_row}"

    aulas_height = max(1, min(len(aulas_catalogo), total_horario_rows))
    aulas_end_row = 3 + aulas_height
    dynamic_aulas_range = f"O4:O{aulas_end_row}"
    aulas_range_cf = f"$O4:O{aulas_end_row}"

    # Armado de datos de la hoja
    horario_data = [["", "", "", "", ""] for _ in range(total_horario_rows)]
    aulas_lateral = list(aulas_catalogo) + [""] * max(0, total_horario_rows - len(aulas_catalogo))

    data = [
        ["Grupo ", group] + [""] * 13,
        [""] * 15,
        ["", "", *DAYS, "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"],
    ]

    for idx in range(total_horario_rows):
        asig_row = normalized_subjects[idx] if idx < len(normalized_subjects) else ["", "", "", "", ""]
        data.append([
            "",
            row_names[idx],
            *horario_data[idx],
            "",
            *asig_row,
            "",
            aulas_lateral[idx],
        ])

    # Fórmulas base
    formulas = [
        {"row": asig_end_row + 1, "col": 12, "value": "Total:"},
        {"row": asig_end_row + 1, "col": 13, "value": f"=COUNTA(I4:I{asig_end_row})"},
        {"row": asig_end_row + 2, "col": 12, "value": "Σ Frec:"},
        {"row": asig_end_row + 2, "col": 13, "value": f"=SUM(K4:K{asig_end_row})"},
        {"row": aulas_end_row + 1, "col": 14, "value": "Total:"},
        {"row": aulas_end_row + 1, "col": 15, "value": f"=COUNTA(O4:O{aulas_end_row})"},
        {"row": horario_end_row + 1, "col": 6, "value": "Ocupados:"},
        {"row": horario_end_row + 1, "col": 7, "value": f"=COUNTA(C4:G{horario_end_row})/{HORARIO_ROW_STEP}"},
    ]

    # Fórmulas por asignatura (solo si hay abreviatura)
    for idx, asig in enumerate(normalized_subjects, start=4):
        abrev = str(asig[0]).strip()
        if not abrev:
            continue
        formulas.append({"row": idx, "col": 13, "value": f"=COUNTIF(C4:G{horario_end_row},I{idx})"})
        formulas.append({"row": idx, "col": 12, "value": f"=K{idx}-M{idx}"})

    # Reglas de formato condicional
    conditional_format_rules = []
    for horario_range in dynamic_horario_ranges:
        conditional_format_rules.extend([
            {
                "tipo": "filas_pares",
                "rango": horario_range,
                "formula": f'AND({{celda}}<>"", COUNTIF({asig_abrev_range},{{celda}})=0)',
                "color": "F4A460",
                "row_step": HORARIO_ROW_STEP,
                "row_start_offset": 0,
            },
            {
                "tipo": "filas_impares",
                "rango": horario_range,
                "formula": f'AND({{celda}}<>"", COUNTIF({aulas_range_cf},{{celda}})=0)',
                "color": "FFD700",
                "row_step": HORARIO_ROW_STEP,
                "row_start_offset": 1,
            },
            {
                "tipo": "pares_con_siguiente",
                "rango": horario_range,
                "formula": 'AND({celda}<>"", {celda_siguiente}="")',
                "color": "FF0000",
                "row_step": HORARIO_ROW_STEP,
                "next_offset": 1,
                "aplicar_a": "siguiente",
            },
        ])

    conditional_format_rules.extend([
        {
            "tipo": "rango",
            "rango": f"J4:J{asig_end_row}",
            "formula": 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            "color": "00FF00",
        },
        {
            "tipo": "rango",
            "rango": f"J4:J{asig_end_row}",
            "formula": 'AND({celda}<>"", M{fila}>0, L{fila}<0)',
            "color": "FF6B6B",
        },
        {
            "tipo": "rango",
            "rango": f"J4:J{asig_end_row}",
            "formula": 'AND({celda}<>"", M{fila}>0, L{fila}>0, L{fila}<K{fila})',
            "color": "FFA500",
        },
    ])

    table_ranges = [
        "B3:G3",
        *dynamic_turnos_ranges,
        dynamic_separator_range,
        *dynamic_horario_ranges,
        "I3:M3",
        dynamic_asig_range,
        dynamic_aulas_range,
    ]

    table_block_sizes = [
        {"range": "B3:G3", "row_step": 1, "col_step": 1},
        {"range": dynamic_turnos_ranges[0], "row_step": HORARIO_ROW_STEP, "col_step": 1},
        {"range": dynamic_turnos_ranges[1], "row_step": HORARIO_ROW_STEP, "col_step": 1},
        {"range": dynamic_separator_range, "row_step": 1, "col_step": 1},
        {"range": dynamic_horario_ranges[0], "row_step": HORARIO_ROW_STEP, "col_step": 1},
        {"range": dynamic_horario_ranges[1], "row_step": HORARIO_ROW_STEP, "col_step": 1},
        {"range": "I3:M3", "row_step": 1, "col_step": 1},
        {"range": dynamic_asig_range, "row_step": 1, "col_step": 1},
        {"range": dynamic_aulas_range, "row_step": 1, "col_step": 1},
    ]

    return {
        "title": group,
        "data": data,
        "column_widths": {i: 14 for i in range(1, 16)},
        "range_styles": [
            {"range": f"I3:I{asig_end_row}", "style": {"bg_color": COLOR_VERDE_ASIGNATURAS}},
            {"range": dynamic_turnos_ranges[0], "style": {"bg_color": COLOR_ROJO_TURNOS}},
            {"range": dynamic_turnos_ranges[1], "style": {"bg_color": COLOR_ROJO_TURNOS}},
        ],
        "table_ranges": table_ranges,
        "horario_data_range": f"C4:G{horario_end_row}",
        "table_block_sizes": table_block_sizes,
        "merge_ranges": build_dynamic_merge_ranges(),
        "table_borders": True,
        "border_color": COLOR_BORDE,
        "border_style": "medium",
        "formulas": formulas,
        "conditional_format_rules": conditional_format_rules,
    }


def build_aulas_data(turnos: int = TURNOS) -> list[list[str]]:
    """Construye la data de la hoja Aulas (vacía, lista para fórmulas)."""
    rows: list[list[str]] = []
    rows.append([""] * 12)
    rows.append(["", ""] + [str(i) for i in range(1, 10)] + ["Lab"])

    turnos_labels = ["1ro", "2do", "3ro", "4to", "5to", "6to"]

    for idx, day in enumerate(DAYS):
        rows.append(["", day, *AULAS_CATALOGO])
        for t in range(turnos):
            rows.append(["", turnos_labels[t], *("" for _ in range(10))])
        if idx < len(DAYS) - 1:
            rows.append([""] * 12)

    return rows


def build_day_blocks(turnos: int = TURNOS) -> list[tuple[int, int, int, str]]:
    """Bloques de días en Aulas: (header_row, row_start, row_end, group_col_letter)."""
    blocks: list[tuple[int, int, int, str]] = []
    row_offset = 3
    for day_idx in range(len(DAYS)):
        header_row = row_offset
        row_start = header_row + 1
        row_end = row_start + turnos - 1
        group_col_letter = col_letter(3 + day_idx)  # C..G en hojas de grupo
        blocks.append((header_row, row_start, row_end, group_col_letter))
        row_offset = row_end + (2 if day_idx < len(DAYS) - 1 else 0)
    return blocks


def build_aulas_formula(groups: list[str], group_cell_ref: str, header_ref: str) -> str:
    """Fórmula que concatena grupos coincidentes para una celda de Aulas."""
    parts = [
        f'IF({group}!{group_cell_ref}={header_ref},{group}!$B$1 & " ","")'
        for group in groups
    ]
    return f'=SUBSTITUTE(TRIM(CONCAT({",".join(parts)}))," ",",")'


def build_aulas_fernando_formulas(
    groups: list[str],
    row_step: int = HORARIO_ROW_STEP,
    turnos: int = TURNOS,
) -> list[dict]:
    """Genera fórmulas cruzadas para la hoja Aulas."""
    if not groups:
        return []

    formulas: list[dict] = []
    aula_offset = 1 if row_step > 1 else 0

    for header_row, row_start, row_end, group_col in build_day_blocks(turnos):
        for aulas_col in range(3, 13):  # C..L
            aulas_col_letter = col_letter(aulas_col)
            header_ref = f"{aulas_col_letter}${header_row}"

            for aulas_row in range(row_start, row_end + 1):
                turno_index = aulas_row - row_start
                turno_offset = 1 if turno_index >= 3 else 0
                group_row = 4 + aula_offset + (turno_index * row_step) + turno_offset
                group_cell_ref = f"${group_col}${group_row}"
                cell = f"{aulas_col_letter}{aulas_row}"

                formulas.append({
                    "cell": cell,
                    "formula": build_aulas_formula(groups, group_cell_ref, header_ref),
                })

    return formulas


def build_aulas_sheet_config(groups: list[str]) -> dict:
    """Construye la hoja Aulas con fórmulas de cruce entre hojas de grupo."""
    data = build_aulas_data()

    table_ranges = ["B3:L9", "B11:L17", "B19:L25", "B27:L33", "B35:L41"]
    range_styles = [
        {"range": "B3:L3", "style": {"bg_color": COLOR_ROJO_AULAS}},
        {"range": "B11:L11", "style": {"bg_color": COLOR_ROJO_AULAS}},
        {"range": "B19:L19", "style": {"bg_color": COLOR_ROJO_AULAS}},
        {"range": "B27:L27", "style": {"bg_color": COLOR_ROJO_AULAS}},
        {"range": "B35:L35", "style": {"bg_color": COLOR_ROJO_AULAS}},
    ]

    return {
        "title": "Aulas",
        "data": data,
        "column_widths": {i: 12 for i in range(1, 13)},
        "range_styles": range_styles,
        "table_ranges": table_ranges,
        "table_borders": True,
        "border_color": "B3B3B3",
        "border_style": "thick",
        "fernando_formulas": build_aulas_fernando_formulas(groups),
    }


def build_config(grupos: list[str], asignaturas_por_grupo: dict[str, list[object]] | None = None) -> dict:
    """Arma config completa para hoja_con_formulas.py."""
    if asignaturas_por_grupo is None:
        asignaturas_por_grupo = {}

    sheets = [
        build_group_sheet_config(grupo, asignaturas_por_grupo.get(grupo, []))
        for grupo in grupos
    ]
    sheets.append(build_aulas_sheet_config(grupos))
    return {"sheets": sheets}


def main() -> None:
    grupos = ["D111", "D211", "C111"]

    # Fase 1: solo nombres de grupos
    config_fase_1 = build_config(grupos)
    generar_excel_personalizado(config_fase_1, "tutorial_fase1_grupos_vacios.xlsx")

    # Fase 2: grupos + asignaturas por grupo
    asignaturas_por_grupo = {
        "D111": [
            ("AL", "Álgebra Lineal", 3),
            ("L", "Lógica", 2),
            ("IP", "Introducción a la Programación", 2),
            ("AM I", "Análisis Matemático I", 2),
        ],
        "D211": [
            ("MA", "Matemática y Aplicaciones", 2),
            ("Prb", "Probabilidades", 2),
            ("BD", "Bases de Datos", 2),
            ("ED", "Estructura de Datos", 2),
        ],
        "C111": [
            ("A I", "Álgebra I", 3),
            ("L", "Lógica", 2),
            ("P", "Programación", 3),
            ("AM I", "Análisis Matemático I", 2),
            ("F", "Filosofía", 2),
        ],
    }

    config_fase_2 = build_config(grupos, asignaturas_por_grupo)
    generar_excel_personalizado(config_fase_2, "tutorial_fase2_con_asignaturas.xlsx")

    print("OK: tutorial_fase1_grupos_vacios.xlsx")
    print("OK: tutorial_fase2_con_asignaturas.xlsx")


if __name__ == "__main__":
    main()
