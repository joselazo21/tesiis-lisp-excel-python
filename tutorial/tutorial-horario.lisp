;; tutorial-horario.lisp
;; Horario de defensas de tesis — junio 2026, Facultad de Matemática y
;; Computación, Universidad de La Habana.
;;
;; Genera Horario_Tesis.xlsx con dos tablas en una sola hoja:
;;
;;   defensas-table   — una fila por defensa (13 filas)
;;                      columnas: dia, hora, estudiante, tutor, presidente,
;;                                secretario, vocal, oponente, local
;;
;;   profesores-table — una fila por profesor (32 filas)
;;                      columnas: nombre, defensas (calculada con COUNTIF)
;;                      formato condicional: resalta en rojo al profesor si
;;                      participa en un slot con dos defensas simultáneas
;;
;; Flujo:
;;   1. def-table         → construye nodos AST (plantillas de tabla)
;;   2. libro / hoja      → ensambla el workbook como árbol AST
;;   3. xl-generate       → compila el árbol a Python (horario-tesis.py)
;;   4. xl-run-generated  → ejecuta el script y crea Horario_Tesis.xlsx

(load (merge-pathnames "dsl-directo.lisp"   *load-truename*))
(load (merge-pathnames "data-tutorial.lisp" *load-truename*))

(format t "~%=== HORARIO DE DEFENSAS DE TESIS — JUNIO 2026 ===~%~%")

;; =====================================================================
;; TABLA 1 — defensas
;; Columnas en el mismo orden que el Excel fuente.
;; Los cinco roles de tribunal (tutor...oponente) son consecutivos para
;; poder referenciarlos como rango con (trange defensas-table tutor oponente).
;; =====================================================================

(def-table defensas-table
  ((dia "") (hora "") (estudiante "") (tutor "")
   (presidente "") (secretario "") (vocal "") (oponente "") (local ""))
  :cell-height 1
  :cell-width  1)

;; =====================================================================
;; TABLA 2 — profesores
;; El nombre incluye el grado (ej. "MSc. Fernando Raúl Rodríguez Flores")
;; para que el COUNTIF localice al profesor en las columnas de tribunal.
;;
;; La columna "defensas" cuenta en cuántos roles aparece el profesor en
;; cualquier defensa de la tabla anterior.
;;
;; El formato condicional usa exists cross-table:
;;
;;   (exists (r :rows-of defensas-table)
;;     :self-in  (tutor presidente secretario vocal oponente)
;;     :matching (dia hora))
;;
;; Semántica: el profesor está en una defensa cuyo slot (dia, hora)
;; tiene más de una defensa programada simultáneamente.
;; El backend genera una FormulaRule SUMPRODUCT que persiste en el .xlsx.
;; =====================================================================

(def-table profesores-table
  ((nombre "") (defensas ""))
  :cell-height 1
  :cell-width  1
  :computed
    ((defensas (countif (trange defensas-table tutor oponente) (col nombre))))
  :render
    (conditional-rendering
      :condition
        (exists (r :rows-of defensas-table)
          :self-in  (tutor presidente secretario vocal oponente)
          :matching (dia hora))
      :target-columns (nombre)))

;; =====================================================================
;; WORKBOOK
;; =====================================================================

(libro horario-tesis
  :filename "Horario_Tesis.xlsx"
  :hojas (list
    (hoja "Defensas"
      (tabla defensas-table   :data *DATOS-DEFENSAS*)
      (tabla profesores-table :data *DATOS-PROFESORES*)
      :fixed-expressions
        (((nav (ultima-fila defensas-table dia) abajo 1)
          (str "Total defensas:"))
         ((nav (ultima-fila defensas-table dia) abajo 1 derecha 2)
          (counta (trange defensas-table estudiante)))
         ((nav (ultima-fila profesores-table nombre) abajo 1)
          (str "Total profesores:"))
         ((nav (ultima-fila profesores-table nombre) abajo 1 derecha 1)
          (sum-range (trange profesores-table defensas)))))))

(xl-generate horario-tesis "horario-tesis.py")
(xl-run-generated "horario-tesis.py")

(format t "~%Hecho: Horario_Tesis.xlsx~%")
