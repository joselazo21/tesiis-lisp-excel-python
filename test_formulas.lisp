; Quick test
(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")
(load "generar_tv_modelo.lisp")
(load "variables_horario_tv.lisp")

(generar-horario-tv :planificacion *tv-planificacion-semanal* :nombre *tv-nombre-canal*)
(quit)