(load "dsl-lenguaje-visual.lisp")
(load "variables_horario_tv.lisp")

(format t "~%=== AST MANUAL (solo macros DSL, cero funciones) ===~%~%")

;; =====================================================================
;; 1. TABLAS
;; =====================================================================

(def-tabla horario-diario (dia hora-inicio-param)
  :columns ((programa "Programa")
            (duracion "Duración (min)")
            (tipo "Tipo")))

(def-tabla resumen-tv ()
  :columns ((dia "Día")
            (programas "Programas")
            (minutos "Minutos")
            (inicio "Inicio")
            (fin "Fin")))

(format t "OK: tablas~%")

;; =====================================================================
;; 2. HOJAS
;; =====================================================================

(def-hoja make-day-sheet (dia hour)
  :name dia
  (horario-diario dia hour
    :data (loop for p in data
                collect (list (getf p :nombre)
                              (getf p :duracion)
                              (getf p :tipo-programa)))))

(def-hoja make-resumen-sheet ()
  :name "Resumen TV"
  (resumen-tv))

(format t "OK: hojas~%")

;; =====================================================================
;; 3. WORKBOOK
;; =====================================================================

(def-horario horario-tv-manual
  :nombre "Horario_TV_Manual.xlsx"
  :hojas (let* ((days '("lunes" "martes" "miercoles" "jueves" "viernes" "sabado" "domingo"))
                (day-names '("Lunes" "Martes" "Miércoles" "Jueves" "Viernes" "Sábado" "Domingo"))
                (day-sheets
                  (loop for key in days
                        for display in day-names
                        for progs = (getf (find-if (lambda (x) (string-equal (getf x :dia) key))
                                                    *tv-planificacion-semanal*)
                                          :programas)
                        for hour = (or (getf (first progs) :hora-inicio) "16:00")
                        collect (make-day-sheet display hour :data progs)))
                (all-progs (loop for d in days
                                 collect (getf (find-if (lambda (x) (string-equal (getf x :dia) d))
                                                        *tv-planificacion-semanal*)
                                              :programas)))
                (total-prog 0)
                (total-min 0)
                (resumen-data
                  (append
                    (loop for progs in all-progs
                          for name in day-names
                          for dur = (loop for p in progs sum (or (getf p :duracion) 0))
                          do (incf total-prog (length progs))
                             (incf total-min dur)
                          collect (list name (length progs) dur
                                        (or (getf (first progs) :hora-inicio) "")
                                        (or (getf (car (last progs)) :hora-final) "")))
                    (list (list "TOTAL" total-prog total-min "" ""))))
                (resumen (make-resumen-sheet :data resumen-data)))
           (append day-sheets (list resumen))))

(xl-generate horario-tv-manual "horario-tv-manual.py")
(xl-run-generated "horario-tv-manual.py")

(format t "~%Hecho: horario-tv-manual generado~%")
