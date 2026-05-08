(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")

(xl-generate 
  (xl-workbook :name "TestExcel"
               :sheets (list (xl-sheet :name "Hoja1"
                                     :tables (list (xl-table :id "mitabla"
                                                          :contenido '(("A" 1) ("B" 2)))))))
  "test_generado.py")