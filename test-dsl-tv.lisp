;; test-dsl-tv.lisp
;; Prueba del DSL de visualización - SIN Excel-ismos visibles
;; El usuario solo usa DSL abstracto; el mapeo a Excel es interno.

(load "dsl-lenguaje-visual.lisp")
(load "variables_horario_tv.lisp")

(format t "~%=== TEST DSL VISUALIZACIÓN TV ===~%~%")

;; =====================================================================
;; 0. Definir expresiones reutilizables (puras, sin contexto de tabla)
;; =====================================================================

(def-expresion calc-hora-final ()
  (_if (non-empty (nombre-del-programa))
    (time-add (hora-inicio) (duracion (nombre-del-programa)))
    (show-nothing)))

(def-expresion calc-hora-inicio ()
  (_if (non-empty (nombre-del-programa))
    (_if (it-is-the-first-row)
      hora-inicio-param
      (hora-final (previous-row (hora-final))))
    (show-nothing)))

(def-expresion highlight-undefined ()
  (_if (is-not-defined (nombre-del-programa) :in programas)
    (set-bg-color :rojo)))

(def-expresion highlight-consecutive-tipo ()
  (_if (is-the-same (tipo (nombre-del-programa))
                     (tipo (previous-row (nombre-del-programa))))
    (set-bg-color :naranja)))

(format t "OK: expresiones definidas~%")

;; =====================================================================
;; 1. Definir tabla horario-diario usando def-tabla
;;    Las fórmulas referencian expresiones por nombre.
;; =====================================================================

(def-tabla horario-diario (dia hora-inicio-param)
  :columns ((hora-inicio "Hora Inicio")
            (hora-final "Hora Final")
            (nombre-del-programa "Nombre"))
  :formula (hora-final :compute calc-hora-final)
  :formula (hora-inicio :compute calc-hora-inicio)
  :cond-format (nombre-del-programa :rule highlight-undefined)
  :cond-format (nombre-del-programa :rule highlight-consecutive-tipo
                :apply-from 2))

(format t "OK: def-tabla horario-diario~%")

;; =====================================================================
;; 2. Definir hoja usando def-hoja (genérico, sin TV hardcodeado)
;; =====================================================================

;; Helpers para construir data y layout específicos de TV
(defun tv-data (day-name programas)
  (let ((hora-inicio (or (getf (first programas) :hora-inicio) "00:00")))
    (append
      (list (list (format nil "Programacion ~a" day-name)
                  "" "" "" "" "" "" "" ""))
      (list (list "" "" "" "" "" "" ""
                  "Hora inicio 1ra fila" hora-inicio))
      (list (list "Programa" "Duracion (min)" "Tipo" ""
                  "Hora de inicio" "Hora de terminacion"
                  "Programa" "Tipo calc" ""))
      (loop for p in programas
            collect (list (getf p :nombre) (getf p :duracion)
                          (getf p :tipo-programa)
                          "" "" "" (getf p :nombre) "" "")))))

(defun tv-layout (data-rows)
  (let ((end-row (+ 3 (1- data-rows))))
    (list
      (list :type :merge :col-start 1 :col-end 9 :row-start 1 :row-end 1)
      (list :type :border :col-start 1 :col-end 3 :row-start 3 :row-end 3)
      (list :type :border :col-start 1 :col-end 3 :row-start 4 :row-end end-row :row-step 2)
      (list :type :border :col-start 5 :col-end 7 :row-start 3 :row-end 3)
      (list :type :border :col-start 5 :col-end 7 :row-start 4 :row-end end-row :row-step 2)
      (list :type :border :col-start 8 :col-end 9 :row-start 2 :row-end 2)
      (list :type :col-width :col 1 :width 42)
      (list :type :col-width :col 2 :width 14)
      (list :type :col-width :col 3 :width 16)
      (list :type :col-width :col 4 :width 3)
      (list :type :col-width :col 5 :width 16)
      (list :type :col-width :col 6 :width 18)
      (list :type :col-width :col 7 :width 42)
      (list :type :col-width :col 8 :width 12)
      (list :type :col-width :col 9 :width 18))))

;; def-hoja genérico: recibe TODO como keyword, no hardcodea nada
(def-hoja planificacion-diaria (nombre)
  :data (tv-data nombre data)
  :layout (tv-layout (length data))
  :params '((hora-inicio-param 9 . 2))
  :ref-table (list :sc 1 :sr 4 :ec 3
                   :er (+ 4 (length data) -1)
                   :cn '(nombre duracion tipo))
  :table-pos '(5 . 4)
  (horario-diario nombre (or (getf (first data) :hora-inicio) "00:00")
    :data (loop for p in data
                collect (list (getf p :nombre)
                              (getf p :duracion)
                              (getf p :tipo-programa)))))

(format t "OK: def-hoja planificacion-diaria~%")

;; =====================================================================
;; 3. Hoja de resumen (hecha con xl-sheet directamente)
;; =====================================================================

(defun construir-resumen-tv (planificacion nombre-canal)
  (let* ((total-prog 0) (total-min 0)
         (data
           (append
             (list (list "Día" "Programas" "Minutos" "Inicio" "Fin"))
             (loop for day in planificacion
                   for dia = (string-capitalize (getf day :dia))
                   for progs = (getf day :programas)
                   for nprog = (length progs)
                   for mins = (loop for p in progs sum (or (getf p :duracion) 0))
                   for inicio = (or (getf (first progs) :hora-inicio) "")
                   for fin = (or (getf (car (last progs)) :hora-final) "")
                   do (incf total-prog nprog) (incf total-min mins)
                   collect (list dia nprog mins inicio fin))
             (list (list "TOTAL" total-prog total-min "" "" "" "" nombre-canal))))
         (max-row (length data)))
    (xl-sheet
      :name "Resumen TV"
      :regions (list
        (xl-region
          :tables (list (xl-table :contenido data))
          :dsl-layout (list
                        (list :type :border
                              :col-start 1 :col-end 7
                              :row-start 1 :row-end max-row)
                        (list :type :col-width :col 1 :width 14)
                        (list :type :col-width :col 2 :width 12)
                        (list :type :col-width :col 3 :width 12)
                        (list :type :col-width :col 4 :width 10)
                        (list :type :col-width :col 5 :width 10)
                        (list :type :col-width :col 6 :width 36)
                        (list :type :col-width :col 7 :width 28)))))))

(format t "OK: constructoras de hojas definidas~%")

;; =====================================================================
;; 4. Construir el libro con def-horario
;; =====================================================================

(let* ((plan *tv-planificacion-semanal*)
       (day-sheets
         (loop for day-config in plan
               for day-name = (string-capitalize (getf day-config :dia))
               for programas = (getf day-config :programas)
                collect (planificacion-diaria day-name
                          :data programas)))
       (resumen (construir-resumen-tv plan *tv-nombre-canal*))
       (all-sheets (append day-sheets (list resumen))))
  (def-horario planificacion-TV
    :nombre "Horario_TV_DSL.xlsx"
    :hojas all-sheets)
  (xl-generate planificacion-TV "planificacion-tv.py")
  (xl-run-generated "planificacion-tv.py"))

(format t "~%=== FIN TEST DSL TV ===~%~%")
