(load "dsl-directo.lisp")
(load "variables_horario_tv.lisp")

(format t "~%=== AST MANUAL DIRECTO ===~%~%")

;; =====================================================================
;; 1. DEFINICIÓN de la tabla de programas (una vez, reutilizable)
;; =====================================================================

(def-table program-table
  ((programa-calc "Programa") (duracion "Duración (min)") (tipo "Tipo")
   (hora-inicio "Hora Inicio") (hora-terminacion "Hora Terminación")
   (tipo-calc "Tipo Calc"))
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
  :render (conditional-rendering
            :condition (_or (equals (previous-of (col tipo))
                                    (col tipo))
                            (equals (col tipo)
                                    (next-of (col tipo))))
            :target-columns (hora-inicio hora-terminacion tipo-calc)))

;; =====================================================================
;; 2. INSTANCIACIÓN — cada día con sus datos
;; =====================================================================

(libro horario-tv-directo
  :filename "Horario_TV_Directo.xlsx"
  :hojas (list
    (hoja "Lunes"
      (tabla program-table
        :data *lunes-progs*
        :params ((hora-inicio-param *lunes-hour*))))
    (hoja "Martes"
      (tabla program-table
        :data *martes-progs*
        :params ((hora-inicio-param *martes-hour*))))
    (hoja "Miércoles"
      (tabla program-table
        :data *miercoles-progs*
        :params ((hora-inicio-param *miercoles-hour*))))
    (hoja "Jueves"
      (tabla program-table
        :data *jueves-progs*
        :params ((hora-inicio-param *jueves-hour*))))
    (hoja "Viernes"
      (tabla program-table
        :data *viernes-progs*
        :params ((hora-inicio-param *viernes-hour*))))
    (hoja "Sábado"
      (tabla program-table
        :data *sabado-progs*
        :params ((hora-inicio-param *sabado-hour*))))
    (hoja "Domingo"
      (tabla program-table
        :data *domingo-progs*
        :params ((hora-inicio-param *domingo-hour*))))
    ;; =================================================================
    ;; Resumen — tabla inline (solo esta, estructura distinta)
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
                  (list (list "TOTAL" total-prog total-min "" ""))))))))
(xl-generate horario-tv-directo "horario-tv-directo.py")
(xl-run-generated "horario-tv-directo.py")

(format t "~%Hecho: horario-tv-directo generado~%")
