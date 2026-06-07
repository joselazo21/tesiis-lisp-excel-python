;; data-tutorial.lisp
;; Datos del horario de defensas de tesis para el tutorial.
;;
;; Profesores: Piad (Dr), Garcia (Dr), Lopez (Msc),
;;             Torres (Msc), Ruiz (Lic), Soto (Msc)
;; Estudiantes: Juan Diaz, Maria Vega, Luis Mora
;; Defensa = tribunal (estudiante + 5 roles) + fecha + hora + local

;; =====================================================================
;; TABLA DE DEFENSAS — una fila por defensa
;; Columnas: estudiante | tutor | oponente | presidente | vocal |
;;           secretario | dia | hora | local
;; =====================================================================

(defparameter *DATOS-DEFENSAS*
  '(("" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "")
    ("Estudiante" "Tutor" "Oponente" "Presidente" "Vocal" "Secretario"
     "Dia" "Hora" "Local")
    ("Juan Diaz"  "Piad"   "Garcia" "Lopez"   "Torres" "Ruiz"   "Lunes"   "10:00" "Aula 1")
    ("Maria Vega" "Garcia" "Lopez"  "Piad"    "Soto"   "Torres" "Lunes"   "14:00" "Aula 2")
    ("Luis Mora"  "Lopez"  "Torres" "Garcia"  "Ruiz"   "Soto"   "Lunes"   "10:00" "Aula 3")))

;; =====================================================================
;; TABLA DE PROFESORES — una fila por profesor
;; La columna "defensas" se calcula en tiempo de compilacion del DSL.
;; =====================================================================

(defparameter *DATOS-PROFESORES*
  '(("" "" "")
    ("" "" "")
    ("Nombre" "Grado" "Defensas")
    ("Piad"   "Dr"    "")
    ("Garcia" "Dr"    "")
    ("Lopez"  "Msc"   "")
    ("Torres" "Msc"   "")
    ("Ruiz"   "Lic"   "")
    ("Soto"   "Msc"   "")))
