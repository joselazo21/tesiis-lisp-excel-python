#!/usr/bin/env python3
"""Convierte un JSON de programacion TV a variables Lisp."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def lisp_string(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{text}"'


def as_int(value: object, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def render_program(program: dict, indent: str = "        ") -> list[str]:
    return [
        f"{indent}(:nombre {lisp_string(program.get('nombre', ''))}",
        f"{indent} :duracion {as_int(program.get('duracion', 0))}",
        f"{indent} :hora-inicio {lisp_string(program.get('hora_inicio', ''))}",
        f"{indent} :hora-final {lisp_string(program.get('hora_final', ''))}",
        f"{indent} :tipo-programa {lisp_string(program.get('tipo_programa', ''))}",
        f"{indent} :tipo-publico {lisp_string(program.get('tipo_publico', ''))})",
    ]


def render_day(day_cfg: dict, indent: str = "  ") -> list[str]:
    lines = [
        f"{indent}(:dia {lisp_string(day_cfg.get('dia', ''))}",
        f"{indent} :programas (",
    ]
    for program in day_cfg.get("programas", []) or []:
        lines.extend(render_program(program, indent=indent + "    "))
    lines.append(f"{indent} ))")
    return lines


def build_lisp_content(json_path: Path, payload: dict) -> str:
    nombre_canal = payload.get("nombre_canal", "Canal TV")
    semanal = payload.get("planificacion_semanal", []) or []

    lines: list[str] = []
    lines.append(";; Archivo auto-generado desde JSON de horario TV")
    lines.append(f";; Fuente: {json_path}")
    lines.append("")
    lines.append(f"(defparameter *tv-json-source* {lisp_string(str(json_path))})")
    lines.append(f"(defparameter *tv-nombre-canal* {lisp_string(nombre_canal)})")
    lines.append("(defparameter *tv-planificacion-semanal*")
    lines.append("  '(")
    for day_cfg in semanal:
        lines.extend(render_day(day_cfg, indent="    "))
    lines.append("   ))")
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convierte JSON TV a variables Lisp")
    parser.add_argument("json_path", help="Ruta del JSON de entrada")
    parser.add_argument(
        "output_lisp",
        nargs="?",
        default="variables_horario_tv.lisp",
        help="Ruta del .lisp de salida (default: variables_horario_tv.lisp)",
    )
    args = parser.parse_args()

    json_path = Path(args.json_path).expanduser().resolve()
    output_lisp = Path(args.output_lisp).expanduser().resolve()

    with json_path.open("r", encoding="utf-8") as f:
        payload = json.load(f)

    content = build_lisp_content(json_path, payload)
    output_lisp.write_text(content, encoding="utf-8")

    print(f"OK: variables Lisp generadas en {output_lisp}")


if __name__ == "__main__":
    main()
