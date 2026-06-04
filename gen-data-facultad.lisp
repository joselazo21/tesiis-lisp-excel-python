;; Genera data-facultad.lisp con datos pre-procesados
;; Produce tres matrices por grupo: turno (6 cols), stats (3 cols), aulas (3 cols)
;; para que ast-facultad.lisp instancie cada tabla con sus datos directos.

(load "variables_horario.lisp")

(defun blank-row (size)
  (make-list size :initial-element ""))

(defun safe-nth (n lst)
  (if (and lst (< n (length lst))) (nth n lst) ""))

(defun pair-horario-rows (data)
  (loop for i from 0 below (length data) by 2
        for asig-row = (nth i data)
        for aula-row = (nth (1+ i) data)
        collect (loop for d from 0 to 4
                      collect (list (safe-nth d asig-row)
                                    (safe-nth d aula-row)))))

(defun build-group-data (grupo-name horario-data asig-data)
  (let* ((days 5)
         (horario-cols days)
         (paired (pair-horario-rows horario-data))
         (turno-labels '("Turno 1" "Turno 2" "Turno 3"
                          "Turno 4" "Turno 5" "Turno 6"))
         ;; --- turno table (6 cols: turno + 5 days) ---
         (turno-title (append (list grupo-name) (blank-row 5)))
         (turno-blank (blank-row 6))
         (turno-header (list "" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes"))
         (turno-rows ())
         ;; --- stats table (5 cols: abrev, asig, frec, faltan, asignadas) ---
         (stats-title (blank-row 5))
         (stats-blank (blank-row 5))
         (stats-header (list "Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas"))
         (stats-rows ())
         ;; --- aulas table (1 col: aulas) ---
         (aulas-title (blank-row 1))
         (aulas-blank (blank-row 1))
         (aulas-header (list "Aulas"))
         (aulas-rows ()))
    (loop for turno-label in turno-labels
          for i from 0
          for pair-row = (or (nth i paired) (blank-row horario-cols))
          for asig-orig = (or (nth i asig-data) '("" "" 0 0 0))
          for abrev = (safe-nth 0 asig-orig)
          for asig-name = (safe-nth 1 asig-orig)
          for frec = (safe-nth 2 asig-orig)
          for aula-list = (loop for day-cell in pair-row
                                for v = (safe-nth 1 day-cell)
                                when (and (stringp v) (> (length (string-trim " " v)) 0))
                                collect v)
          do
          (push (append (list turno-label) pair-row) turno-rows)
          (push (list abrev asig-name (princ-to-string frec) "0" "") stats-rows)
          (dolist (aula aula-list)
            (push (list aula) aulas-rows)))
    (values (append (list turno-title turno-blank turno-header) (nreverse turno-rows))
            (append (list stats-title stats-blank stats-header) (nreverse stats-rows))
            (append (list aulas-title aulas-blank aulas-header) (nreverse aulas-rows)))))

(defun write-row (row stream)
  (format stream "    (")
  (loop for cell in row
        for i from 0
        do (when (> i 0) (format stream " "))
           (format stream "~s" cell))
  (format stream ")~%"))

;; Grupos
(defparameter *grupos*
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

;; Generar data-facultad.lisp
(with-open-file (out "data-facultad.lisp"
                     :direction :output
                     :if-exists :supersede)
  (format out ";; data-facultad.lisp — Datos pre-procesados para horario de facultad~2%")

  ;; Data por grupo — 3 matrices separadas (turno, stats, aulas)
  (dolist (g *grupos*)
    (let* ((name (symbol-name (first g)))
           (horario (symbol-value (second g)))
           (asig (symbol-value (third g))))
      (multiple-value-bind (turno-data stats-data aulas-data)
          (build-group-data name horario asig)
        ;; Turno data (6 cols)
        (let ((var (intern (concatenate 'string "*DATOS-TURNO-" name "*"))))
          (format out "(defparameter ~a~%  '(" var)
          (dolist (row turno-data) (write-row row out))
          (format out "  ))~2%"))
        ;; Stats data (3 cols)
        (let ((var (intern (concatenate 'string "*DATOS-STATS-" name "*"))))
          (format out "(defparameter ~a~%  '(" var)
          (dolist (row stats-data) (write-row row out))
          (format out "  ))~2%"))
        ;; Aulas data (3 cols)
        (let ((var (intern (concatenate 'string "*DATOS-AULAS-" name "*"))))
          (format out "(defparameter ~a~%  '(" var)
          (dolist (row aulas-data) (write-row row out))
          (format out "  ))~2%")))))

  ;; Aulas data (hoja "Aulas") - separado por día
  (let* ((dias '("Lunes" "Martes" "Miercoles" "Jueves" "Viernes"))
         (aulas-titles '("Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5"
                         "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab"))
         (total-cols 11)
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
              ("" "" "" "" "" "" "" "" "" "")))))
    (loop for dia in dias
          for day-data in aulas-rows
          do
          (let ((rows nil))
            (push (append (list dia) aulas-titles) rows)
            (loop for turno-row in day-data
                  for turno-num from 1
                  do (push (append (list (princ-to-string turno-num)) turno-row) rows))
            (setf rows (nreverse rows))
            (let ((var (intern (concatenate 'string "*DATOS-AULAS-" (string-upcase dia) "*"))))
              (format out "(defparameter ~a~%  '(" var)
              (dolist (row rows) (write-row row out))
              (format out "  ))~2%"))))))

(format t "data-facultad.lisp generated~%")
