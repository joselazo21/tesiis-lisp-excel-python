import json

C_SHEETS = ["C111", "C121", "C122", "C211", "C212", "C311", "C312", "C411", "C412"]
D_SHEETS = ["D111", "D211", "D311", "D411"]
M_SHEETS = ["M111", "M211", "M311", "M411"]
ALL_SHEETS = C_SHEETS + D_SHEETS + M_SHEETS

DAYS_CONFIG = {
    "Lunes":       {"col": "C", "sheets": ALL_SHEETS, "base_row": 4},
    "Martes":      {"col": "D", "sheets": ALL_SHEETS, "base_row": 12},
    "Miércoles":   {"col": "E", "sheets": ALL_SHEETS, "base_row": 20},
    "Jueves":      {"col": "F", "sheets": ALL_SHEETS, "base_row": 28},
    "Viernes":     {"col": "G", "sheets": ALL_SHEETS, "base_row": 36},
}

SLOT_ROWS = {
    "1ro": 5,
    "2do": 7,
    "3ro": 9,
}

AULA_COLS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]

def generate_formula(day_col, group_row, sheets, aula_col):
    parts = []
    for sheet in sheets:
        parts.append(f'SI({sheet}!${day_col.upper()}${group_row}={aula_col}$2,{sheet}!$B$1 & " "; "")')
    inner = "; ".join(parts)
    return f'=SUSTITUIR(ESPACIOS(CONCATENAR({inner})); " "; ",")'

aulas_formulas = []

for day_name, cfg in DAYS_CONFIG.items():
    day_col = cfg["col"]
    sheets = cfg["sheets"]
    base_row = cfg["base_row"]
    for slot_idx, (slot_name, group_row) in enumerate(SLOT_ROWS.items()):
        aulas_row = base_row + slot_idx
        for aula_col in AULA_COLS:
            formula = generate_formula(day_col, group_row, sheets, aula_col)
            aulas_formulas.append({
                "cell": f"{aula_col}{aulas_row}",
                "excel": formula,
            })

data = {"Aulas": aulas_formulas}

with open("formulas_fernando_adaptadas_es.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

print(f"Generadas {len(aulas_formulas)} fórmulas para hoja Aulas")
print(f"Hojas por fórmula: {len(ALL_SHEETS)} ({', '.join(ALL_SHEETS)})")
print(f"Días: {list(DAYS_CONFIG.keys())}")
