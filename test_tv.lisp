(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")

;; Test xl-workbook creation (with keyword like before)
(let ((sheet (xl-sheet :name "Lunes" 
                     :tables (list (xl-table :id "tv_lunes" 
                                          :contenido '(("08:00" "09:00" 60 "Noticias" "Programa" "Público")))))))
  (print "Sheet created OK")
  (let ((wb (xl-workbook :name "Test" :sheets (list sheet))))
    (print "Workbook created OK")
    (xl-generate wb "tv_test.py")
    (format t "Generado: tv_test.py~%")))