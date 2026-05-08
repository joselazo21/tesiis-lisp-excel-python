(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")
(load "generar_tv_modelo.lisp")
(load "variables_horario_tv.lisp")

(format t "~%Generando horario TV con modelo~%")
(generar-horario-tv :planificacion *tv-planificacion-semanal*
                  :nombre-canal *tv-nombre-canal*
                  :output-file "tv_test.py")
(format t "~%Listo~%")