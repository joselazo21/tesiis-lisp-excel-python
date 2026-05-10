; ast_manual.lisp - Manual AST for TV schedule
(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")
(load "variables_horario_tv.lisp")

; Get programas for a day
(defun get-prog (dia)
  (getf (find-if #'(lambda (d) (string-equal (getf d :dia) dia)) *tv-planificacion-semanal*) :programas))

; First hour
(defun first-hour (prog-list)
  (or (getf (first prog-list) :hora-inicio) "16:00"))

; Sheet name
(defun sheet-name (n)
  (cond ((= n 1) "01-Lunes") ((= n 2) "02-Martes") ((= n 3) "03-Miércoles")
        ((= n 4) "04-Jueves") ((= n 5) "05-Viernes") ((= n 6) "06-Sábado")
        ((= n 7) "07-Domingo")))

; Data for one day - FIXED
(defun make-data (prog hour dia)
  (let ((d nil))
    (setf d (list (list (concatenate 'string "Programacion " dia) nil nil nil nil nil nil nil nil)))
    (setf d (cons (list nil nil nil nil nil nil nil "Hora inicio 1ra fila" hour) d))
    (setf d (cons (list "Programa" "Duracion (min)" "Tipo" nil "Hora de inicio" "Hora de terminacion" "Programa" "Tipo calc" nil) d))
    (dolist (p prog)
      (setf d (cons (list (getf p :nombre) (getf p :duracion) (getf p :tipo-programa) nil "" "" (getf p :nombre) "") d)))
    (reverse d)))

; Formulas for one day - FIX proper list building
(defun make-formulas (last)
  (let ((f nil))
    ; Header row 3
    (push (make-instance 'clase-xl-formula :row 3 :col 5 :value "=IF(G3<>\"\",$I$2)") f)
    (push (make-instance 'clase-xl-formula :row 3 :col 6 :value "=IF(G3<>\"\",TEXT(TIMEVALUE(E3)+IFERROR(VLOOKUP(G3,$A$4:$C$19,2,FALSE),0)/1440,\"hh:mm\"))") f)
    (push (make-instance 'clase-xl-formula :row 3 :col 8 :value "=IF(G3<>\"\",IFERROR(VLOOKUP(G3,$A$4:$C$19,3,FALSE),\"\"))") f)
    ; Data rows
    (loop for r from 4 to last do
      (push (make-instance 'clase-xl-formula :row r :col 5 :value (concatenate 'string "=IF(G" (write-to-string r) "<>\"\",F" (write-to-string (1- r)) ")")) f)
      (push (make-instance 'clase-xl-formula :row r :col 6 :value (concatenate 'string "=IF(G" (write-to-string r) "<>\"\",TEXT(TIMEVALUE(E" (write-to-string r) ")+IFERROR(VLOOKUP(G" (write-to-string r) ",$A$4:$C$" (write-to-string last) ",2,FALSE),0)/1440,\"hh:mm\"))")) f)
      (push (make-instance 'clase-xl-formula :row r :col 8 :value (concatenate 'string "=IF(G" (write-to-string r) "<>\"\",IFERROR(VLOOKUP(G" (write-to-string r) ",$A$4:$C$" (write-to-string last) ",3,FALSE),\"\"))")) f))
    (reverse f)))

; Make day sheet - WITH formulas list node
(defun make-sheet (dia n)
  (let* ((prog (get-prog dia)) (hour (first-hour prog)) (nprog (length prog))
         (data (make-data prog hour dia))
         (nrows (+ 3 nprog))
         (formulas (make-formulas nrows))
         (formula-list (make-instance 'clase-xl-formula-list :items formulas)))
    (format t "~a: ~a formulas=~a~%" dia nprog (length formulas))
    (xl-sheet :name (sheet-name n) :tables (list (xl-table :id dia :contenido data))
              :formulas formula-list
              :table-borders t :border-color "B7B7B7" :border-style "thin"
              :table-ranges (list (concatenate 'string "A3:C" (write-to-string nrows))
                                (concatenate 'string "E3:G" (write-to-string nrows)) "H2:I2")
              :column-widths (xl-column-widths :pairs '(1 42 2 14 3 16 4 3 5 16 6 18 7 42 8 12 9 18))
              :header-style (xl-header-style :bold t :align "center" :bg-color "D9EAD3"))))

; Sum durations
(defun sum-dur (prog)
  (let ((s 0)) (dolist (p prog) (setf s (+ s (or (getf p :duracion) 0)))) s))

; Resumen sheet - match target
(defun make-resumen ()
  (let* ((p1 (get-prog "lunes")) (p2 (get-prog "martes")) (p3 (get-prog "miercoles"))
         (p4 (get-prog "jueves")) (p5 (get-prog "viernes")) (p6 (get-prog "sabado"))
         (p7 (get-prog "domingo"))
         (total-p (+ (length p1) (length p2) (length p3) (length p4) (length p5) (length p6) (length p7)))
         (total-m (+ (sum-dur p1) (sum-dur p2) (sum-dur p3) (sum-dur p4) (sum-dur p5) (sum-dur p6) (sum-dur p7)))
         (d nil))
    ; Headers
    (setf d (list (list "Día" "Programas" "Minutos" "Inicio" "Fin" "Tipos" "Públicos")))
    ; Days: Lunes first, then to Domingo
    (setf d (cons (list "Lunes" (length p1) (sum-dur p1) nil nil nil nil nil) d))
    (setf d (cons (list "Martes" (length p2) (sum-dur p2) nil nil nil nil nil) d))
    (setf d (cons (list "Miércoles" (length p3) (sum-dur p3) nil nil nil nil nil) d))
    (setf d (cons (list "Jueves" (length p4) (sum-dur p4) nil nil nil nil nil) d))
    (setf d (cons (list "Viernes" (length p5) (sum-dur p5) nil nil nil nil nil) d))
    (setf d (cons (list "Sábado" (length p6) (sum-dur p6) nil nil nil nil nil) d))
    (setf d (cons (list "Domingo" (length p7) (sum-dur p7) nil nil nil nil nil) d))
    ; TOTAL at end
    (setf d (cons (list "TOTAL" total-p total-m nil nil nil nil "Canal Habana") d))
    (setf d (reverse d))
    (format t "Resumen: ~a ~a~%" total-p total-m)
    (xl-sheet :name "Resumen TV" :tables (list (xl-table :id "resumen" :contenido d))
              :table-borders t :border-color "B7B7B7" :border-style "thin"
              :column-widths (xl-column-widths :pairs '(1 14 2 12 3 12 4 10 5 10 6 36 7 28))
              :header-style (xl-header-style :bold t :align "center" :bg-color "F4CCCC"))))

; Create workbook - fixed
(let* ((s1 (make-sheet "Lunes" 1)) (s2 (make-sheet "Martes" 2)) (s3 (make-sheet "Miercoles" 3))
       (s4 (make-sheet "Jueves" 4)) (s5 (make-sheet "Viernes" 5)) (s6 (make-sheet "Sabado" 6))
       (s7 (make-sheet "Domingo" 7)) (sr (make-resumen))
       (sheets (list s1 s2 s3 s4 s5 s6 s7 sr)))
  (defparameter *wb* (xl-workbook :name "Canal Habana.xlsx" :sheets sheets))
  (format t "~%Done: ~a sheets~%" (length sheets)))

; Generate
(format t "Generating...~%")
(xl-generate *wb* "horario_tv.py")
(format t "Generated: horario_tv.py~%")
