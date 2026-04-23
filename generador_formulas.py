"""Generador de fórmulas cruzadas para la hoja Aulas.

Este módulo contiene las funciones para generar fórmulas que cruzan
entre hojas de grupos y la hoja Aulas.
"""


def col_letter(col_num: int) -> str:
    """Convierte número de columna (1-based) a letra Excel."""
    value = ""
    n = col_num
    while n > 0:
        n, rem = divmod(n - 1, 26)
        value = chr(65 + rem) + value
    return value


def build_day_blocks(dias: list[str], turnos: int) -> list[tuple[int, int, int, str]]:
    """Bloques de días en Aulas: (header_row, row_start, row_end, group_col_letter)."""
    blocks = []
    row_offset = 1
    for day_idx in range(len(dias)):
        header_row = row_offset
        row_start = header_row + 1
        row_end = row_start + turnos - 1
        group_col_letter = col_letter(3 + day_idx)  # C..G en hojas de grupo
        blocks.append((header_row, row_start, row_end, group_col_letter))
        row_offset = row_end + (2 if day_idx < len(dias) - 1 else 0)
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
    dias: list[str],
    row_step: int = 3,
    turnos: int = 6,
) -> list[dict]:
    """Genera fórmulas cruzadas para la hoja Aulas."""
    if not groups:
        return []

    formulas = []
    aula_offset = 1 if row_step > 1 else 0

    for header_row, row_start, row_end, group_col in build_day_blocks(dias, turnos):
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
