;; =====================================================================
;; ast_manual.lisp - TV schedule using ONLY DSL macros + make-instance
;;
;; Zero helper functions (defun). Todo se construye inline.
;; Sin layout, estilo, colocación ni tamaño — eso es del generate-code.
;; Solo datos y estructura.
;; =====================================================================

(load "dsl-lenguaje-visual.lisp")
(load "variables_horario_tv.lisp")

(format t "~%=== AST MANUAL (datos + formulas, sin estilo) ===~%~%")

;; =====================================================================
;; 1. EXPRESIONES
;; =====================================================================

(def-expresion calc-hora-inicio ()
  (_if (non-empty (programa-calc))
    (_if (it-is-the-first-row)
      hora-inicio-param
      (hora-terminacion (previous-row (hora-terminacion))))
    (show-nothing)))

(def-expresion calc-hora-terminacion ()
  (_if (non-empty (programa-calc))
    (time-add (hora-inicio) (duracion (programa-calc)))
    (show-nothing)))

(def-expresion calc-tipo-calc ()
  (_if (non-empty (programa-calc))
    (tipo (programa-calc))
    (show-nothing)))

(format t "OK: expresiones~%")

;; =====================================================================
;; 2. TABLA
;; =====================================================================

(def-tabla horario-diario (dia hora-inicio-param)
  :columns ((hora-inicio "Hora Inicio")
            (hora-terminacion "Hora Terminación")
            (programa-calc "Programa")
            (tipo-calc "Tipo Calc"))
  :formula (hora-inicio :compute calc-hora-inicio)
  :formula (hora-terminacion :compute calc-hora-terminacion)
  :formula (tipo-calc :compute calc-tipo-calc))

(format t "OK: def-tabla~%")

;; =====================================================================
;; 3. HOJAS DE DÍA
;;
;; Sin :layout — el backend elige bordes, anchos, merge por defecto.
;; =====================================================================

(def-hoja make-day-sheet (dia hour)
  :name dia
  :data (append
          (list (list (concatenate 'string "Programacion " dia)
                      "" "" "" "" "" "" "" ""))
          (list (list "" "" "" "" "" "" "" "Hora inicio 1ra fila" hour))
          (list (list "Programa" "Duracion (min)" "Tipo" ""
                      "Hora de inicio" "Hora de terminacion"
                      "Programa" "Tipo calc" ""))
          (loop for p in data
                collect (list (getf p :nombre) (getf p :duracion)
                              (getf p :tipo-programa)
                              "" "" "" (getf p :nombre) "" "")))
  :params '((hora-inicio-param 9 . 2))
  :ref-table (list :sc 1 :sr 4 :ec 3
                   :er (+ 3 (length data))
                   :cn '(nombre duracion tipo))
  :table-pos '(5 . 4)
  :fernando-formulas (list (make-instance 'clase-xl-fernando-formula
                            :cell (format nil "I~a" (+ 3 (length data)))
                            :formula "=1+1"))
  (horario-diario dia hour
    :data (loop for p in data
                collect (list (getf p :nombre)
                              (getf p :duracion)
                              (getf p :tipo-programa)))))

(format t "OK: def-hoja make-day-sheet~%")

;; =====================================================================
;; 4. HOJA DE RESUMEN
;;
;; Sin :dsl-layout. Solo datos.
;; =====================================================================

(let* ((days '("lunes" "martes" "miercoles" "jueves" "viernes" "sabado" "domingo"))
       (day-names '("Lunes" "Martes" "Miércoles" "Jueves" "Viernes" "Sábado" "Domingo"))
       (all-progs (loop for d in days
                        collect (getf (find-if (lambda (x)
                                                 (string-equal (getf x :dia) d))
                                               *tv-planificacion-semanal*)
                                     :programas)))
       (total-prog 0)
       (total-min 0)
       (data (append
               (list (list "Día" "Programas" "Minutos" "Inicio" "Fin" "Tipos" "Públicos"))
               (loop for progs in all-progs
                     for name in day-names
                     for dur = (loop for p in progs sum (or (getf p :duracion) 0))
                     do (incf total-prog (length progs))
                        (incf total-min dur)
                     collect (list name (length progs) dur
                                   (or (getf (first progs) :hora-inicio) "")
                                   (or (getf (car (last progs)) :hora-final) "")
                                   "" ""))
               (list (list "TOTAL" total-prog total-min "" "" "" "" *tv-nombre-canal*)))))
  (defparameter resumen-sheet
    (xl-sheet :name "Resumen TV"
      :regions (list
        (xl-region
          :tables (list (xl-table :contenido data)))))))

(format t "OK: resumen~%")

;; =====================================================================
;; 5. WORKBOOK
;; =====================================================================

(let* ((day-names '("Lunes" "Martes" "Miércoles" "Jueves" "Viernes" "Sábado" "Domingo"))
       (day-keys '("lunes" "martes" "miercoles" "jueves" "viernes" "sabado" "domingo"))
       (day-sheets
         (loop for key in day-keys
               for display in day-names
               for progs = (getf (find-if (lambda (x) (string-equal (getf x :dia) key))
                                          *tv-planificacion-semanal*)
                                 :programas)
               for hour = (or (getf (first progs) :hora-inicio) "16:00")
               collect (make-day-sheet display hour :data progs)))
       (all (append day-sheets (list resumen-sheet))))
  (def-horario horario-tv-manual
    :nombre "Horario_TV_Manual.xlsx"
    :hojas all)
  (xl-generate horario-tv-manual "horario-tv-manual.py")
  (xl-run-generated "horario-tv-manual.py"))

(format t "~%Hecho: horario-tv-manual generado~%")
