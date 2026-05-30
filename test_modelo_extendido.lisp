(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")

(let* ((tbl (xl-table :id "test_data" 
                  :contenido '(("Program A" 30) ("Program B" 45))
                  :headers '("Programa" "Duracion")))
       (sheet (xl-sheet :name "Test"
                    :regions (list
                      (xl-region
                        :tables (list tbl))))
       (wb (xl-workbook :name "Test.xlsx"
                      :sheets (list sheet))))
  (xl-generate wb "test_modelo.py")
  (format t "~%Generacion completada~%"))