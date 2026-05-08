(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")
(load "generar_tv_modelo.lisp")

(defparameter *tv-test*
  '((:dia "lunes"
     :programas ((:nombre "Test Program"
              :duracion 30
              :hora-inicio "16:00"
              :hora-final "16:30"
              :tipo-programa "informativo"
              :tipo-publico "adulto")))))

(format t "~%Generando TV extendido~%")
(generar-horario-tv :planificacion *tv-test*
                  :nombre-canal "Test"
                  :output-file "test_tv_ext.py")
(format t "~%Listo~%")