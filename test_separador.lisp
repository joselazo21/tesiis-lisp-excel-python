(load "codigo-tesis.lisp")
(load "variables_horario.lisp")
(load "replicar_propuesta_ods.lisp")

(defparameter *test-data* 
  (list 
    (list "A" "B") (list "a" "b") 
    (list "C" "D") (list "c" "d")
    (list "E" "F") (list "e" "f")
    (list "G" "H") (list "g" "h")
    (list "I" "J") (list "i" "j")
    (list "K" "L") (list "k" "l")))

(let* ((horario-row-step 3)
       (cantidad-turnos 6)
       (horario-data-normalizado (normalizar-horario-data *test-data* cantidad-turnos horario-row-step))
       (separador-index (* 3 horario-row-step))
       (base-row-names (nombres-filas-por-turnos cantidad-turnos horario-row-step))
       (blank-horario-row (fila-vacia-como (first horario-data-normalizado)))
       (horario-row-names (insert-at-index base-row-names separador-index ""))
       (horario-data-con-separador (insert-at-index horario-data-normalizado separador-index blank-horario-row)))
       
  (format t "Normalizado length: ~a~%" (length horario-data-normalizado))
  (format t "Row names length: ~a~%" (length horario-row-names))
  (format t "Data con separador length: ~a~%" (length horario-data-con-separador))
  (format t "Row names: ~a~%" horario-row-names))
