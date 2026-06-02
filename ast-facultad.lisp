;; ast-facultad.lisp
;; Intento de modelar el horario de facultad con el DSL/AST actual.
;; Identifica qué se puede expresar y qué falta.
;;
;; OBJETIVO: usar SOLO las clases AST existentes y macros del DSL directo.
;; Sin modificar ast-def.lisp ni dsl-directo.lisp.

(load "dsl-directo.lisp")
(load "variables_horario.lisp")

;; =====================================================================
;; GAP 1: No hay concepto de "fila separadora" ni "merged cells" en xl-table
;; GAP 2: No hay "title row" — solo headers (1 fila) y data (N filas)
;; GAP 3: Las fórmulas computed son siempre por-columna × fila.
;;         No hay fórmulas en posiciones fijas (totales, etiquetas).
;; GAP 4: No hay "conditional_format_rules" con placeholders.
;;         style-rules solo hace highlight por igualdad de valores, no COUNTIF.
;; GAP 5: No hay "fernando_formulas" — fórmulas cross-sheet.
;; GAP 6: No hay "merge_ranges", "table_block_sizes", "table_ranges".
;; =====================================================================

;; =====================================================================
;; UTILITIES
;; =====================================================================

(defun blank-row (size)
  (make-list size :initial-element ""))

(defun safe-nth (n lst)
  (if (and lst (< n (length lst))) (nth n lst) ""))

(defun pair-horario-rows (data)
  "Agrupa las 12 filas de horario (2 por turno) en 6 registros
   con 10 columnas: (lun-subj lun-class mar-subj mar-class ...)."
  (loop for i from 0 below (length data) by 2
        for asig-row = (nth i data)
        for aula-row = (nth (1+ i) data)
        collect (loop for d from 0 to 4
                      append (list (safe-nth d asig-row)
                                   (safe-nth d aula-row)))))

(defun build-group-data (grupo-name horario-data asig-data)
  "Construye las filas completas para una hoja de grupo.
   Layout (21 columnas):
     A: vacío, B: turno-label
     C-L: horario emparejado (lun-s lun-c mar-s mar-c ... vie-s vie-c) = 10 cols
     M: gap, N: abrev, O: asig-name, P: frec, Q: faltan (raw), R: asignadas (raw)
     S: gap, T: aulas, U: align"
  (let* ((days 5) (cols-per-day 2) (horario-cols (* days cols-per-day))
         (paired (pair-horario-rows horario-data))
         (turno-labels '("Turno 1" "Turno 2" "Turno 3"
                         "Turno 4" "Turno 5" "Turno 6"))
         (suffix-cols 9)  ;; M-U = gap, abrev, name, frec, faltan, asignadas, gap, aulas, align
         (total-cols (+ 2 horario-cols suffix-cols))
         (title-row (append (list "Grupo " grupo-name)
                            (blank-row (- total-cols 2))))
         (blank (blank-row total-cols))
         (header-row (list "" ""
                           "Lunes" "" "Martes" "" "Miércoles" "" "Jueves" "" "Viernes" ""
                           ""
                           "Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas"
                           "" "Aulas" ""))
         (data-rows ()))
    (loop for turno-label in turno-labels
          for i from 0
          for pair-row = (or (nth i paired) (blank-row horario-cols))
          for asig-orig = (or (nth i asig-data) '("" "" 0 0 0))
          for abrev = (safe-nth 0 asig-orig)
          for asig-name = (safe-nth 1 asig-orig)
          for frec = (safe-nth 2 asig-orig)
          for aula-list = (loop for j from 1 below (length pair-row) by 2
                                for v = (nth j pair-row)
                                when (and (stringp v) (> (length (string-trim " " v)) 0))
                                collect v)
          for aula = (if aula-list (format nil "~{~a~^,~}" aula-list) "")
          do (push (append (list "" turno-label)  ; A, B
                           pair-row               ; C-L
                           (list ""               ; M gap
                                 abrev            ; N
                                 asig-name        ; O
                                 (princ-to-string frec) ; P
                                 "0"              ; Q faltan (raw)
                                 ""               ; R asignadas (raw)
                                 ""               ; S gap
                                 aula             ; T aulas
                                 ""))             ; U align
                    data-rows))
    (setf data-rows (nreverse data-rows))
    (list title-row blank header-row data-rows)))

;; =====================================================================
;; DEFINIR LA TABLA DE GRUPO — ahora con fórmulas reales
;; =====================================================================

(defun make-group-table (grupo-name horario-data asig-data)
  (destructuring-bind (title blank header data-rows)
      (build-group-data grupo-name horario-data asig-data)
    (let* ((data (append (list title blank header) data-rows))
           (num-data-rows (length data-rows))
           (first-data-row 4)
           (last-data-row (+ first-data-row num-data-rows -1))
           (horario-range "$C${first-row}:$L${last-row}")
           (col-names '(_a _b
                        lun-s lun-c mar-s mar-c
                        mie-s mie-c jue-s jue-c
                        vie-s vie-c
                        _gap-prefix
                        abrev asig frec faltan asignadas
                        _gap-suffix aulas _align))
           ;; Per-row computed formulas:
           ;;   asignadas (col 18) = COUNTIF(C$4:L$9, N{row})
           ;;   faltan (col 17) = frec - asignadas
            (computeds
              `((asignadas . ,(countif horario-range (col abrev)))
                (faltan . ,(subtract (col frec) (col asignadas)))))
           ;; Fixed formulas (totals below data):
           ;;   Total asignaturas: row = last+1, col 14 = COUNTA(N4:N9)
           ;;   Σ Frec: row = last+2, col 13 = SUM(P4:P9)
           (total-row (+ last-data-row 1))
           (sum-row (+ last-data-row 2))
           (fixed-fs
             (list
               (fixed-formula total-row 14 (counta "N${first-row}:N${last-row}"))
               (fixed-formula total-row 13 (str "Total:"))
               (fixed-formula sum-row 13 (sum-range "P${first-row}:P${last-row}"))
               (fixed-formula sum-row 12 (str "Σ Frec:")))))
      (list data col-names computeds fixed-fs first-data-row last-data-row))))

;; =====================================================================
;; INSTANCIAR PARA TODOS LOS GRUPOS
;; =====================================================================

(defparameter *grupos-facultad*
  '((D111 *horario-d111* *asignaturas-d111*)
    (D211 *horario-d211* *asignaturas-d211*)
    (D311 *horario-d311* *asignaturas-d311*)
    (D411 *horario-d411* *asignaturas-d411*)
    (C111 *horario-c111* *asignaturas-c111*)
    (C121 *horario-c121* *asignaturas-c121*)
    (C122 *horario-c122* *asignaturas-c122*)
    (C211 *horario-c211* *asignaturas-c211*)
    (C212 *horario-c212* *asignaturas-c212*)
    (C311 *horario-c311* *asignaturas-c311*)
    (C312 *horario-c312* *asignaturas-c312*)
    (C411 *horario-c411* *asignaturas-c411*)
    (C412 *horario-c412* *asignaturas-c412*)
    (M111 *horario-m111* *asignaturas-m111*)
    (M211 *horario-m211* *asignaturas-m211*)
    (M311 *horario-m311* *asignaturas-m311*)
    (M411 *horario-m411* *asignaturas-m411*)))

;; Por cada grupo creamos el xl-table (contenido + fórmulas + estilo)
(defun build-all-group-tables ()
  (loop for (name horario-var asig-var) in *grupos-facultad*
        for grupo-name = (symbol-name name)
        for (data col-names computeds fixed-fs first-data-row last-data-row)
          = (make-group-table grupo-name
                              (symbol-value horario-var)
                              (symbol-value asig-var))
        collect
          (list
            grupo-name
            (xl-table
              :contenido data
              :headers nil
              :col-names col-names
              :computed computeds
              :fixed-formulas fixed-fs
              :first-row first-data-row
              :params nil
              :style-rules nil)
            first-data-row
            last-data-row)))

;; =====================================================================
;; HOJA AULAS — cruce cross-sheet
;; =====================================================================
;; GAP mayor: las celdas de la hoja Aulas necesitan fórmulas
;; como: =IF(C111!$C$5=C$3, C111!$B$1 & " ", "")
;; que referencian otras hojas. El AST actual no tiene
;; "cross-sheet references" ni "fernando_formulas".
;;
;; Construimos los datos planos pero sin las fórmulas.
;; Documentamos que hay que inyectarlas vía fernando_formulas.

(defun build-aulas-data ()
  (let* ((dias '("Lunes" "Martes" "Miércoles" "Jueves" "Viernes"))
         (aulas-titles '("1" "2" "3" "4" "5" "6" "7" "8" "9" "Lab"))
         (aulas-rows
           '((("" "" "" "" "" "C111,C112" "" "" "" "C113")
              ("" "" "" "" "C112" "C113" "" "" "" "")
              ("" "" "" "" "" "C111,C113" "" "C112" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" ""))
             (("" "C112" "" "" "" "C111" "C113" "" "" "")
              ("C111" "" "" "" "" "C112,C113" "" "" "" "")
              ("" "" "" "" "" "C111,C112,C113" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" ""))
             (("" "" "" "" "C112" "C111" "" "C113" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "C112" "" "" "C113" "C111" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" ""))
             (("C113" "" "" "" "C112" "C111" "" "" "" "")
              ("C113" "" "" "" "C112" "C111" "" "" "" "")
              ("C113" "" "" "" "C112" "C111" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" ""))
             (("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "C112")
              ("" "" "" "" "C111" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" "")
              ("" "" "" "" "" "" "" "" "" ""))))
         (rows (list (blank-row 13)            ;; row 1 blank
                     (append (list "" "")      ;; row 2 header
                             aulas-titles
                             (list "")))))     ;; align to 13
    ;; Agregar cada día con su contenido
    (loop for dia in dias
          for day-data in aulas-rows
          for i from 0
          do (push (append (list "" dia) aulas-titles (list "")) rows)
             (loop for turno-row in day-data
                   do (push (append (list "" (princ-to-string (1+ (position turno-row day-data :test #'eq))))
                                    turno-row
                                    (list ""))
                            rows))
             (unless (= i (1- (length dias)))
               (push (blank-row 13) rows)))
    (nreverse rows)))

;; =====================================================================
;; CONSTRUIR EL WORKBOOK
;; =====================================================================

(format t "~%=== AST FACULTAD ===~%~%")
(format t "Construyendo ~a grupos...~%" (length *grupos-facultad*))

(let ((group-tables (build-all-group-tables))
      (aulas-data (build-aulas-data))
      (sheets '()))
  (format t "~&Resumen por grupo:~%")
  ;; Hojas de grupo
  (dolist (entry group-tables)
    (destructuring-bind (name tbl first-row last-row) entry
      (let ((data-len (length (contenido tbl))))
        (push (xl-sheet :name name
                        :regions (list (xl-region :tables (list tbl))))
              sheets)
        (format t "  ~a: ~a filas, first=~a last=~a~%"
                name data-len first-row last-row))))

  ;; Hoja Aulas
  (push
    (xl-sheet :name "Aulas"
              :regions (list
                         (xl-region
                           :tables
                           (list
                             (xl-table
                               :contenido aulas-data
                               :headers nil
                               :col-names '(_a _b aula1 aula2 aula3 aula4 aula5
                                            aula6 aula7 aula8 aula9 lab _align)
                               :computed nil
                               :params nil
                               :style-rules nil)))))
    sheets)

  (setf sheets (nreverse sheets))
  (format t "  Aulas: ~a filas~%" (length aulas-data))

  (let ((wb (xl-workbook :name "Horario_Facultad.xlsx"
                         :sheets sheets)))
    (defparameter *libro-facultad* wb)
    (format t "~%Libro FACULTAD creado con ~a hojas~%" (length sheets))
    (xl-generate wb "horario-facultad.py")
    (xl-run-generated "horario-facultad.py")))

;; =====================================================================
;; DIAGNÓSTICO: Lo que falta en el AST/DSL actual
;; =====================================================================
(format t "~2%========================================================~%")
(format t "  DIAGNÓSTICO: Capacidades faltantes del AST/DSL~%")
(format t "========================================================~%~%")

(format t "1. TITLE ROW~%")
(format t "   xl-table tiene :headers (1 fila) pero no título/descripción~%")
(format t "   antes de los headers. El título 'Grupo X' se mete como~%")
(format t "   primera fila de :contenido, perdiendo semántica.~%~%")

(format t "2. SEPARATOR ROW~%")
(format t "   Entre Turno 3 y Turno 4 hay una fila separadora vacía.~%")
(format t "   No hay slot para filas de separación.~%~%")

(format t "3. MERGE RANGES~%")
(format t "   Cada turno-label ('Turno 1') debería fusionarse en 2 filas.~%")
(format t "   xl-table no tiene :merge-ranges. El backend lo soporta~%")
(format t "   (hoja_con_formulas.py acepta merge_ranges) pero el AST~%")
(format t "   no lo genera.~%~%")

(format t "4. FÓRMULAS EN POSICIONES FIJAS~%")
(format t "   :computed genera fórmulas por-columna × fila. El horario~%")
(format t "   necesita fórmulas en celdas específicas: Total, Σ Frec,~%")
(format t "   Ocupados, etc. Además COUNTIF requiere un rango fijo~%")
(format t "   ($C$4:$G$15) no variable por fila.~%~%")

(format t "5. CONDITIONAL FORMATTING RULES~%")
(format t "   El backend (hoja_con_formulas.py) soporta~%")
(format t "   conditional_format_rules con placeholders~%")
(format t "   ({celda}, {celda_siguiente}, COUNTIF, etc.) pero el AST~%")
(format t "   solo tiene style-rules por comparación de valores.~%")
(format t "   Faltan: 'filas_pares', 'filas_impares', 'pares_con_siguiente'~%~%")

(format t "6. CROSS-SHEET FORMULAS (Aulas)~%")
(format t "   La hoja Aulas necesita referencias como C111!$C$5.~%")
(format t "   compile-excel-formula solo genera fórmulas intra-hoja.~%")
(format t "   No hay support para 'fernando_formulas'.~%~%")

(format t "7. TABLE BLOCK BORDERS~%")
(format t "   El horario necesita bordes por bloques de 2 filas (turnos).~%")
(format t "   xl-table no tiene :table-ranges ni :table-block-sizes.~%")
(format t "   El slot :border-color existe pero solo un color global.~%~%")

(format t "8. PARAM SLOT vs FIRST-ROW~%")
(format t "   first-row se calcula automáticamente con params=0 → 2~%")
(format t "   En facultad tenemos 3 filas prefijo (title+blank+header)~%")
(format t "   → first-row debería ser 4. No hay control explícito.~%~%")

(format t "========================================================~%")
(format t "  Backend (hoja_con_formulas.py) ya soporta todo lo~%")
(format t "  anterior (merge_ranges, conditional_format_rules,~%")
(format t "  fernando_formulas, table_block_sizes). Falta que el~%")
(format t "  AST lo modele y generate-code lo serialice.~%")
(format t "========================================================~%")

(provide "ast-facultad")
