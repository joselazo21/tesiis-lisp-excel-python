#!/usr/bin/env bash
# run.sh — ejecuta el tutorial básico
#
# Uso:
#   chmod +x run.sh
#   ./run.sh
#
# Lo que hace:
#   1. Carga tutorial-basico.lisp con SBCL.
#   2. SBCL expande macros, construye el AST y genera tutorial-basico.py.
#   3. SBCL invoca python3 para ejecutar ese script y crear Tutorial_Basico.xlsx.

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

SALIDA="Tutorial_Basico.xlsx"

echo "=== Tutorial Básico: Tabla, Promedio y Resaltado ==="
echo ""
echo "Paso 1/2 — Cargando DSL y generando script Python..."

sbcl --noinform --script tutorial-basico.lisp 2>&1 \
  | grep -v "^$" \
  | grep -v "^STYLE-WARNING" \
  | grep -v "^WARNING:"

echo ""
echo "Paso 2/2 — Verificando resultado..."

if [ -f "$SALIDA" ]; then
    echo "OK: $SALIDA creado correctamente."
    ls -lh "$SALIDA"
else
    echo "ERROR: no se encontró $SALIDA." >&2
    exit 1
fi

echo ""
echo "=== Listo ==="
