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
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *lunes-progs*
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
        :params ((hora-inicio-param *lunes-hour*))))

    ;; =================================================================
    ;; Martes
    ;; =================================================================
    (hoja "Martes"
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *martes-progs*
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
        :params ((hora-inicio-param *martes-hour*))))

    ;; =================================================================
    ;; Miércoles
    ;; =================================================================
    (hoja "Miércoles"
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *miercoles-progs*
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
        :params ((hora-inicio-param *miercoles-hour*))))

    ;; =================================================================
    ;; Jueves
    ;; =================================================================
    (hoja "Jueves"
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *jueves-progs*
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
        :params ((hora-inicio-param *jueves-hour*))))

    ;; =================================================================
    ;; Viernes
    ;; =================================================================
    (hoja "Viernes"
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *viernes-progs*
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
        :params ((hora-inicio-param *viernes-hour*))))

    ;; =================================================================
    ;; Sábado
    ;; =================================================================
    (hoja "Sábado"
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *sabado-progs*
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
        :params ((hora-inicio-param *sabado-hour*))))

    ;; =================================================================
    ;; Domingo
    ;; =================================================================
    (hoja "Domingo"
      (tabla ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
              (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
              (tipo-calc "Tipo Calc"))
        :data *domingo-progs*
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
        :params ((hora-inicio-param *domingo-hour*))))

    ;; =================================================================
    ;; Resumen
    ;; =================================================================
    (hoja "Resumen TV"
      (tabla ((dia "Día") (programas "Programas") (minutos "Minutos")
              (inicio "Inicio") (fin "Fin"))
        :data (let* ((total-prog 0)
                     (total-min 0))
                (append
                  (loop for entry in *semana-resumen*
                        for name = (first entry)
                        for _hour = (second entry)
                        for _progs = (third entry)
                        for dur = (loop for p in _progs sum (or (second p) 0))
                        do (incf total-prog (length _progs))
                           (incf total-min dur)
                        collect (list name (length _progs) dur
                                      (or _hour "")
                                      ""))
                  (list (list "TOTAL" total-prog total-min "" "")))))))
)
(xl-generate horario-tv-directo "horario-tv-directo.py")
(xl-run-generated "horario-tv-directo.py")

(format t "~%Hecho: horario-tv-directo generado~%")
