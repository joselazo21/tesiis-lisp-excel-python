;; tutorial-horario.lisp
;; Horario de defensas de tesis — programa DSL completo.
;;
;; Genera Horario_Tesis.xlsx con dos tablas en una sola hoja:
;;
;;   defensas-table   — una fila por defensa
;;                      columnas: estudiante, tutor, oponente, presidente,
;;                                vocal, secretario, dia, hora, local
;;
;;   profesores-table — una fila por profesor
;;                      columnas: nombre, grado, defensas (calculada con COUNTIF)
;;                      formato condicional: resalta en rojo al profesor si
;;                      participa en alguna defensa con conflicto de horario
;;
;; Flujo de ejecución:
;;   1. def-table         → construye nodos AST (plantillas reutilizables de tabla)
;;   2. libro / hoja      → ensambla el workbook como árbol de nodos AST
;;   3. xl-generate       → compila el árbol a Python (horario-tesis.py)
;;   4. xl-run-generated  → ejecuta el script y crea Horario_Tesis.xlsx

(load "dsl-directo.lisp")
(load "data-tutorial.lisp")

(format t "~%=== HORARIO DE DEFENSAS DE TESIS ===~%~%")

;; =====================================================================
;; TABLA 1 — defensas
;; Nueve columnas: cinco de tribunal (tutor...secretario) más dia, hora,
;; local.  Las cinco columnas de rol son consecutivas para poder
;; referenciarlas como rango con (trange defensas-table tutor secretario).
;; No tiene fórmulas computed ni reglas de coloreado propias; el análisis
;; de solapamiento se expresa desde profesores-table.
;; =====================================================================

(def-table defensas-table
  ((estudiante "") (tutor "") (oponente "") (presidente "")
   (vocal "") (secretario "") (dia "") (hora "") (local ""))
  :cell-height 1
  :cell-width 1)

;; =====================================================================
;; TABLA 2 — profesores
;; La columna "defensas" se calcula con COUNTIF sobre el rango de roles
;; (tutor...secretario) de todas las filas de defensas-table.
;;
;; El formato condicional usa el cuantificador exists cross-table:
;;
;;   (exists (r :rows-of defensas-table)
;;     :self-in  (tutor oponente presidente vocal secretario)
;;     :matching (dia hora))
;;
;; Semántica: "existe alguna fila de defensas-table en la que el nombre
;; de este profesor aparece en un rol, y cuyo slot (dia, hora) tiene más
;; de una defensa programada".
;;
;; El backend detecta que la condición es un nodo xl-expr-exists con
;; dominio cross-table y genera automáticamente una FormulaRule con la
;; fórmula SUMPRODUCT+COUNTIFS en el .xlsx.  El coloreado persiste y se
;; recalcula en Excel sin necesidad de volver a ejecutar el DSL.
;; =====================================================================

(def-table profesores-table
  ((nombre "") (grado "") (defensas ""))
  :cell-height 1
  :cell-width 1
  :computed
    ((defensas (countif (trange defensas-table tutor secretario) (col nombre))))
  :render
    (conditional-rendering
      :condition
        (exists (r :rows-of defensas-table)
          :self-in  (tutor oponente presidente vocal secretario)
          :matching (dia hora))
      :target-columns (nombre)))

;; =====================================================================
;; WORKBOOK — una sola hoja con ambas tablas en la misma región horizontal.
;; Las expresiones fijas se colocan mediante nav, que calcula la posición
;; relativa a la última fila de cada tabla; si se añaden más filas de datos
;; los totales se desplazan automáticamente sin cambiar el DSL.
;; =====================================================================

(libro horario-tesis
  :filename "Horario_Tesis.xlsx"
  :hojas (list
    (hoja "Defensas"
      (tabla defensas-table   :data *DATOS-DEFENSAS*)
      (tabla profesores-table :data *DATOS-PROFESORES*)
      :fixed-expressions
        (((nav (ultima-fila defensas-table estudiante) abajo 1)
          (str "Total defensas:"))
         ((nav (ultima-fila defensas-table estudiante) abajo 1 derecha 1)
          (counta (trange defensas-table estudiante)))
         ((nav (ultima-fila profesores-table nombre) abajo 1)
          (str "Total profesores:"))
         ((nav (ultima-fila profesores-table nombre) abajo 1 derecha 2)
          (sum-range (trange profesores-table defensas)))))))

(xl-generate horario-tesis "horario-tesis.py")
(xl-run-generated "horario-tesis.py")

(format t "~%Hecho: Horario_Tesis.xlsx~%")
