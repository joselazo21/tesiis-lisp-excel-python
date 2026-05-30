#!/bin/bash
cd /home/jose/Proyectos/Scripting/python_excel_tesis
echo "=== Generating TV Schedule (Directo) ==="
echo "1. Running SBCL (AST directo)..."
sbcl --noinform --eval '(load "ast-manual-directo.lisp")' --quit 2>&1 | grep -v "^$" | grep -v "STYLE-WARNING" | grep -v "^WARNING:" | grep -v "TABLA is being" | grep -v "HOJA is being" | grep -v "LIBRO is being"

if [ -f "horario-tv-directo.py" ]; then
    echo "2. Python generated: horario-tv-directo.py"
    echo "3. Running Python..."
    python3 horario-tv-directo.py
    if [ -f "Horario_TV_Directo.xlsx" ]; then
        echo "=== DONE ==="
        ls -lh "Horario_TV_Directo.xlsx"
    else
        echo "ERROR: Excel not created"
    fi
else
    echo "ERROR: Python not generated"
fi
