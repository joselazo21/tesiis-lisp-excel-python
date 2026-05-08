(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")
(load "generar_tv_modelo.lisp")
(load "variables_horario_tv.lisp")

(format t "~%Generando horario TV completo~%")
(generar-horario-tv :planificacion *tv-planificacion-semanal*
                  :nombre-canal *tv-nombre-canal*
                  :output-file "horario_tv.py")
(format t "~%Listo~%")