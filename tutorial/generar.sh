#!/usr/bin/env bash
# generar.sh — ejecuta el flujo completo del tutorial del DSL
#
# Uso:
#   chmod +x generar.sh   (solo la primera vez)
#   ./generar.sh
#
# El script debe ejecutarse desde el directorio donde reside:
#   cd /ruta/al/tutorial
#   ./generar.sh
#
# Lo que hace:
#   1. Carga el programa DSL con SBCL.
#   2. SBCL expande las macros, construye el AST y genera horario-tesis.py.
#   3. SBCL invoca python3 para ejecutar ese script y crear Horario_Tesis.xlsx.

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

SALIDA="Horario_Tesis.xlsx"

echo "=== Tutorial: Horario de Defensas de Tesis ==="
echo ""
echo "Paso 1/2 — Cargando DSL y generando script Python..."

sbcl --noinform --script tutorial-horario.lisp 2>&1 \
  | grep -v "^$" \
  | grep -v "^STYLE-WARNING" \
  | grep -v "^WARNING:"

echo ""
echo "Paso 2/2 — Verificando resultado..."

if [ -f "$SALIDA" ]; then
    echo "OK: $SALIDA creado correctamente."
    ls -lh "$SALIDA"
else
    echo "ERROR: no se encontro $SALIDA." >&2
    exit 1
fi

echo ""
echo "=== Listo ==="
