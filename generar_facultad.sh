#!/bin/bash
cd /home/jose/Proyectos/Scripting/python_excel_tesis
echo "=== Generating Facultad Schedule ==="
echo "1. Generating data..."
sbcl --noinform --script gen-data-facultad.lisp 2>&1 | grep -v "^$" | grep -v "^STYLE-WARNING" | grep -v "^WARNING:"
echo "2. Building workbook and generating Excel..."
sbcl --noinform --script ast-facultad.lisp 2>&1 | grep -v "^$" | grep -v "^STYLE-WARNING" | grep -v "^WARNING:"
if [ -f "Horario_Facultad.xlsx" ]; then
    echo "=== DONE ==="
    ls -lh "Horario_Facultad.xlsx"
else
    echo "ERROR: Excel not created"
fi
