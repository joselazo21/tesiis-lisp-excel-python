#!/bin/bash
cd /home/jose/Proyectos/Scripting/python_excel_tesis
echo "=== Generating Tutorial Schedule (Defensas de Tesis) ==="
echo "1. Building workbook and generating Excel..."
sbcl --noinform --script tutorial-horario.lisp 2>&1 | grep -v "^$" | grep -v "^STYLE-WARNING" | grep -v "^WARNING:"
if [ -f "Horario_Tesis.xlsx" ]; then
    echo "=== DONE ==="
    ls -lh "Horario_Tesis.xlsx"
else
    echo "ERROR: Excel not created"
fi
