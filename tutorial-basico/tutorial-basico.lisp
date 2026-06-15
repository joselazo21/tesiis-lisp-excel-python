;; tutorial-basico.lisp
;; Tutorial básico: tabla simple con columnas, fórmula y resaltado.
;;
;; Genera Tutorial_Basico.xlsx con una tabla de 4 columnas:
;;
;;   nombre    — nombre de la persona
;;   nota1     — primera nota (número)
;;   nota2     — segunda nota (número)
;;   promedio  — promedio de nota1 y nota2 (calculado con fórmula)
;;
;; Además hay un valor fijo (nota-minima) definido como parámetro.
;; Las personas cuyo promedio supera nota-minima se resaltan en rojo.
;;
;; Flujo:
;;   1. def-table   → plantilla con columna calculada + formato condicional
;;   2. libro / hoja → ensambla el workbook
;;   3. xl-generate → genera el script Python
;;   4. xl-run-generated → ejecuta el script y crea el .xlsx

(load (merge-pathnames "dsl-directo.lisp" *load-truename*))

(format t "~%=== TUTORIAL BÁSICO ===~%~%")

;; ─────────────────────────────────────────────────────────────────────
;; DATOS
;; ─────────────────────────────────────────────────────────────────────

(defparameter *DATOS*
  '(("" "" "" "")
    ("" "" "" "")
    ("Nombre" "Nota 1" "Nota 2" "Promedio")
    ("Ana"     "85" "90" "")
    ("Luis"    "70" "75" "")
    ("Elena"   "92" "88" "")
    ("Carlos"  "60" "65" "")
    ("Sofía"   "78" "82" "")
    ("Pedro"   "95" "91" "")))

;; ─────────────────────────────────────────────────────────────────────
;; TABLA
;; ─────────────────────────────────────────────────────────────────────

(def-table tabla-notas
  ((nombre "") (nota1 "") (nota2 "") (promedio ""))
  :cell-height 1
  :cell-width  1
  :computed
    ((promedio (promedio (col nota1) (col nota2))))
  :render
    (conditional-rendering
      :condition (gt (col promedio) (param nota-minima))
      :target-columns (nombre nota1)))

;; ─────────────────────────────────────────────────────────────────────
;; LIBRO
;; ─────────────────────────────────────────────────────────────────────

(libro tutorial-basico
  :hojas (list
    (hoja "Notas"
      (tabla tabla-notas
        :data *DATOS*
        :params ((nota-minima 80))))))

(xl-generate tutorial-basico "tutorial-basico.py")
(xl-run-generated "tutorial-basico.py")

(format t "~%Hecho: Archivo-Excel.xlsx~%")
