#!/bin/bash
cd /home/jose/Proyectos/Scripting/python_excel_tesis
echo "=== Generating TV Schedule ==="
echo "1. Running clisp (AST manual)..."
clisp -q -q <<'EOF' 2>&1 | grep -v "^$" | grep -v "^ADVERTENCIA"
(load "ast_manual.lisp")
(quit)
EOF

if [ -f "horario_tv.py" ]; then
    echo "2. Python generated: horario_tv.py"
    echo "3. Running Python..."
    python3 horario_tv.py
    if [ -f "Canal Habana.xlsx" ]; then
        echo "=== DONE ==="
        ls -lh "Canal Habana.xlsx"
    else
        echo "ERROR: Excel not created"
    fi
else
    echo "ERROR: Python not generated"
fi
