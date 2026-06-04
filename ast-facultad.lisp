;; ast-facultad.lisp — Horario de facultad con SOLO macros del DSL
;; Tres tablas separadas (turno+días, estadísticas, aulas) dentro
;; de una misma región por hoja. La región añade columnas separadoras
;; entre los tres bloques.

(load "dsl-directo.lisp")
(load "data-facultad.lisp")

(format t "~%=== AST FACULTAD ===~%~%")

;; =====================================================================
;; TABLA 1 — turno + días de la semana
;; =====================================================================

(def-table turno-table
  ((turno "") (lun "") (mar "") (mie "") (jue "") (vie ""))
  :first-row 4
  :cell-height 2
  :cell-width 1
  :paired-columns (lun mar mie jue vie)
  :fixed-formulas (((cell 16 7) (str "Total:"))
                   ((cell 17 7) (sum-range (range frec)))
                   ((cell 17 6) (str "Σ Frec:"))))

;; =====================================================================
;; TABLA 2 — estadísticas (abreviatura, asignatura, frecuencia)
;; =====================================================================

(def-table stats-table
  ((abrev "") (asig "") (frec "") (faltan "") (asignadas ""))
  :first-row 4
  :cell-height 1
  :cell-width 1
  :computed ((asignadas (countif (range lun vie) (col abrev)))
             (faltan (subtract (col frec) (col asignadas))))
  :fixed-formulas (((cell 16 1) (counta (range abrev)))))

;; =====================================================================
;; TABLA 3 — aulas por grupo (solo contenido)
;; =====================================================================

(def-table aulas-table
  ((aulas ""))
  :first-row 4
  :cell-height 1
  :cell-width 1)

;; =====================================================================
;; TABLA — aulas por día (hoja "Aulas")
;; Una sola declaración; :dia se pasa como inst-param al instanciar.
;; =====================================================================

(defparameter *grupos-all*
  '("D111" "D211" "D311" "D411"
    "C111" "C121" "C122" "C211" "C212" "C311" "C312" "C411" "C412"
    "M111" "M211" "M311" "M411"))

(def-table aulas-dia-table
  ((dia "") (aula1 "") (aula2 "") (aula3 "") (aula4 "") (aula5 "")
   (aula6 "") (aula7 "") (aula8 "") (aula9 "") (lab ""))
  :first-row 2 :cell-height 1 :cell-width 1
  :inst-params (dia)
  :computed
    ((aula1 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 1"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula2 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 2"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula3 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 3"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula4 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 4"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula5 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 5"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula6 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 6"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula7 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 7"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula8 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 8"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (aula9 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 9"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))
     (lab   (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Lab"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))))

;; =====================================================================
;; CONSTRUIR EL WORKBOOK
;; =====================================================================

(libro horario-facultad
  :filename "Horario_Facultad.xlsx"
  :hojas (list
    (hoja "D111"
      (tabla turno-table :data *DATOS-TURNO-D111*)
      (tabla stats-table :data *DATOS-STATS-D111*)
      (tabla aulas-table :data *DATOS-AULAS-D111*))
    (hoja "D211"
      (tabla turno-table :data *DATOS-TURNO-D211*)
      (tabla stats-table :data *DATOS-STATS-D211*)
      (tabla aulas-table :data *DATOS-AULAS-D211*))
    (hoja "D311"
      (tabla turno-table :data *DATOS-TURNO-D311*)
      (tabla stats-table :data *DATOS-STATS-D311*)
      (tabla aulas-table :data *DATOS-AULAS-D311*))
    (hoja "D411"
      (tabla turno-table :data *DATOS-TURNO-D411*)
      (tabla stats-table :data *DATOS-STATS-D411*)
      (tabla aulas-table :data *DATOS-AULAS-D411*))
    (hoja "C111"
      (tabla turno-table :data *DATOS-TURNO-C111*)
      (tabla stats-table :data *DATOS-STATS-C111*)
      (tabla aulas-table :data *DATOS-AULAS-C111*))
    (hoja "C121"
      (tabla turno-table :data *DATOS-TURNO-C121*)
      (tabla stats-table :data *DATOS-STATS-C121*)
      (tabla aulas-table :data *DATOS-AULAS-C121*))
    (hoja "C122"
      (tabla turno-table :data *DATOS-TURNO-C122*)
      (tabla stats-table :data *DATOS-STATS-C122*)
      (tabla aulas-table :data *DATOS-AULAS-C122*))
    (hoja "C211"
      (tabla turno-table :data *DATOS-TURNO-C211*)
      (tabla stats-table :data *DATOS-STATS-C211*)
      (tabla aulas-table :data *DATOS-AULAS-C211*))
    (hoja "C212"
      (tabla turno-table :data *DATOS-TURNO-C212*)
      (tabla stats-table :data *DATOS-STATS-C212*)
      (tabla aulas-table :data *DATOS-AULAS-C212*))
    (hoja "C311"
      (tabla turno-table :data *DATOS-TURNO-C311*)
      (tabla stats-table :data *DATOS-STATS-C311*)
      (tabla aulas-table :data *DATOS-AULAS-C311*))
    (hoja "C312"
      (tabla turno-table :data *DATOS-TURNO-C312*)
      (tabla stats-table :data *DATOS-STATS-C312*)
      (tabla aulas-table :data *DATOS-AULAS-C312*))
    (hoja "C411"
      (tabla turno-table :data *DATOS-TURNO-C411*)
      (tabla stats-table :data *DATOS-STATS-C411*)
      (tabla aulas-table :data *DATOS-AULAS-C411*))
    (hoja "C412"
      (tabla turno-table :data *DATOS-TURNO-C412*)
      (tabla stats-table :data *DATOS-STATS-C412*)
      (tabla aulas-table :data *DATOS-AULAS-C412*))
    (hoja "M111"
      (tabla turno-table :data *DATOS-TURNO-M111*)
      (tabla stats-table :data *DATOS-STATS-M111*)
      (tabla aulas-table :data *DATOS-AULAS-M111*))
    (hoja "M211"
      (tabla turno-table :data *DATOS-TURNO-M211*)
      (tabla stats-table :data *DATOS-STATS-M211*)
      (tabla aulas-table :data *DATOS-AULAS-M211*))
    (hoja "M311"
      (tabla turno-table :data *DATOS-TURNO-M311*)
      (tabla stats-table :data *DATOS-STATS-M311*)
      (tabla aulas-table :data *DATOS-AULAS-M311*))
    (hoja "M411"
      (tabla turno-table :data *DATOS-TURNO-M411*)
      (tabla stats-table :data *DATOS-STATS-M411*)
      (tabla aulas-table :data *DATOS-AULAS-M411*))
    (hoja-v "Aulas"
      (tabla aulas-dia-table :dia lun :data *DATOS-AULAS-LUNES*)
      (tabla aulas-dia-table :dia mar :data *DATOS-AULAS-MARTES*)
      (tabla aulas-dia-table :dia mie :data *DATOS-AULAS-MIERCOLES*)
      (tabla aulas-dia-table :dia jue :data *DATOS-AULAS-JUEVES*)
      (tabla aulas-dia-table :dia vie :data *DATOS-AULAS-VIERNES*))))

(xl-generate horario-facultad "horario-facultad.py")
(xl-run-generated "horario-facultad.py")

(format t "~%Hecho: Horario_Facultad.xlsx~%")
