#!/bin/bash
# Simple test - generate 2 day schedule

cd /home/jose/Proyectos/Scripting/python_excel_tesis

echo "=== Testing TV Model ==="

# Create a minimal test
sbcl --noinform --non-interactive --eval '
(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")
(load "generar_tv_modelo.lisp")

; Define test data
(defparameter *test-data* (list
  (list :dia "lunes" :programas (list 
    (list :nombre "Noticias" :duracion 30 :hora-inicio "08:00" :hora-final "08:30" :tipo-programa "informativo" :tipo-publico "adulto")
    (list :nombre "Deportes" :duracion 60 :hora-inicio "08:30" :hora-final "09:30" :tipo-programa "deportivo" :tipo-publico "adulto")))
  (list :dia "martes" :programas (list 
    (list :nombre "Cine" :duracion 90 :hora-inicio "10:00" :hora-final "11:30" :tipo-programa "cine" :tipo-publico "adulto")))))

; Generate
(generar-horario-tv :planificacion *test-data* :nombre-canal "Test" :output-file "tv_test.py")
(quit)
' 2>&1 | tail -3

# Run Python
if [ -f "tv_test.py" ]; then
    echo "Running Python..."
    python3 tv_test.py
    ls -lh tv_test.*
fi