(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")

(let* ((header-st (xl-header-style :bold t :bg-color "D9EAD3"))
       (tbl (xl-table :id "test_data" 
                  :contenido '(("Program A" 30) ("Program B" 45))
                  :headers '("Programa" "Duracion")))
       (sheet (xl-sheet :name "Test"
                    :tables (list tbl)
                    :table-borders t
                    :border-color "B7B7B7"
                    :border-style "thin"
                    :header-style header-st))
       (wb (xl-workbook :name "Test.xlsx"
                      :sheets (list sheet))))
  (xl-generate wb "test_modelo.py")
  (format t "~%Generacion completada~%"))