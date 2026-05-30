(load "dsl-directo.lisp")
(load "variables_horario_tv.lisp")

(format t "~%=== AST MANUAL DIRECTO (sin plantillas, instancias explicitas) ===~%~%")

(libro horario-tv-directo
  :filename "Horario_TV_Directo.xlsx"
  :hojas (list
    ;; =================================================================
    ;; Lunes
    ;; =================================================================
    (hoja "Lunes"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "lunes"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Martes
    ;; =================================================================
    (hoja "Martes"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "martes"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Miércoles
    ;; =================================================================
    (hoja "Miércoles"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "miercoles"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Jueves
    ;; =================================================================
    (hoja "Jueves"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "jueves"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Viernes
    ;; =================================================================
    (hoja "Viernes"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "viernes"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Sábado
    ;; =================================================================
    (hoja "Sábado"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "sabado"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Domingo
    ;; =================================================================
    (hoja "Domingo"
      (let* ((_data (find-if (lambda (x) (string-equal (getf x :dia) "domingo"))
                             *tv-planificacion-semanal*))
             (_progs (getf _data :programas))
             (_first-hour (or (getf (first _progs) :hora-inicio) "16:00")))
        (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
                (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
                (tipo-calc "Tipo Calc"))
          :data (loop for p in _progs
                      collect (list (getf p :nombre)
                                    (getf p :duracion)
                                    (getf p :tipo-programa)))
          :computed ((hora-inicio (_if (non-empty (col programa-calc))
                                     (_if (it-is-the-first-row)
                                       (param hora-inicio-param)
                                       (col hora-terminacion (previous-row (col hora-terminacion))))
                                     (show-nothing)))
                     (hora-terminacion (_if (non-empty (col programa-calc))
                                          (time-add (col hora-inicio) (lookup (col programa-calc) duracion))
                                          (show-nothing)))
                     (tipo-calc (_if (non-empty (col programa-calc))
                                   (lookup (col programa-calc) tipo)
                                   (show-nothing))))
          :params ((hora-inicio-param _first-hour)))))

    ;; =================================================================
    ;; Resumen
    ;; =================================================================
    (hoja "Resumen TV"
      (tabla ((dia "Día") (programas "Programas") (minutos "Minutos")
              (inicio "Inicio") (fin "Fin"))
        :data (let* ((days '("lunes" "martes" "miercoles" "jueves" "viernes" "sabado" "domingo"))
                     (day-names '("Lunes" "Martes" "Miércoles" "Jueves" "Viernes" "Sábado" "Domingo"))
                     (all-progs (mapcar (lambda (d)
                                          (getf (find-if (lambda (x) (string-equal (getf x :dia) d))
                                                          *tv-planificacion-semanal*)
                                                :programas))
                                        days))
                     (total-prog 0)
                     (total-min 0))
                (append
                  (loop for progs in all-progs
                        for name in day-names
                        for dur = (loop for p in progs sum (or (getf p :duracion) 0))
                        do (incf total-prog (length progs))
                           (incf total-min dur)
                        collect (list name (length progs) dur
                                      (or (getf (first progs) :hora-inicio) "")
                                      (or (getf (car (last progs)) :hora-final) "")))
                  (list (list "TOTAL" total-prog total-min "" "")))))))
)
(xl-generate horario-tv-directo "horario-tv-directo.py")
(xl-run-generated "horario-tv-directo.py")

(format t "~%Hecho: horario-tv-directo generado~%")
