;; data-facultad.lisp — Datos pre-procesados para horario de facultad

(defparameter *DATOS-TURNO-D111*
  '(    ("D111" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("F" "Aula 8") ("ICD" "Aula 7") ("AL" "Aula 7*") ("AM I"
                                                                 "Aula 7*") ("AL"
                                                                             "Aula 7*"))
    ("Turno 2" ("L" "Aula 6*") ("AL" "Aula 6*") ("EF" "SEDER") ("L" "Aula 7") (""
                                                                               ""))
    ("Turno 3" ("IP" "Aula 7") ("AM I" "Aula 6*") ("AM I" "Aula 7*") ("IP"
                                                                      "Lab") (""
                                                                              ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-D111*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("AL" "Álgebra Lineal" "3" "0" "")
    ("L" "Lógica" "2" "0" "")
    ("IP" "Introducción a la Programación" "2" "0" "")
    ("AM I" "Análisis Matemático I" "2" "0" "")
    ("ICD" "Introducción a la Ciencia de Datos" "2" "0" "")
    ("F" "Filosofía" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-D111*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 8")
    ("Aula 7")
    ("Aula 7*")
    ("Aula 7*")
    ("Aula 7*")
    ("Aula 6*")
    ("Aula 6*")
    ("SEDER")
    ("Aula 7")
    ("Aula 7")
    ("Aula 6*")
    ("Aula 7*")
    ("Lab")
  ))

(defparameter *DATOS-TURNO-D211*
  '(    ("D211" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 4" ("ED" "c 7") ("BD" "c 7") ("MA (EDO) cp 6 (con" "C211") ("BD"
                                                                        "cp 7") (""
                                                                                 ""))
    ("Turno 5" ("VD" "c 7") ("EP" "c 7") ("EP" "c 7") ("Prb" "cp 7") ("" ""))
    ("Turno 6" ("MA" "c 6") ("Prb" "c 7") ("EF 4:45pm a 5:35pm" "") ("ED"
                                                                     "cp 7") (""
                                                                              ""))
  ))

(defparameter *DATOS-STATS-D211*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("MA" "Matemática y Aplicaciones" "2" "0" "")
    ("Prb" "Probabilidades" "2" "0" "")
    ("BD" "Bases de Datos" "2" "0" "")
    ("ED" "Estructura de Datos" "2" "0" "")
    ("VD" "Visualización de Datos" "2" "0" "")
    ("EP" "Economía Política" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-D211*
  '(    ("")
    ("")
    ("Aulas")
    ("c 7")
    ("c 7")
    ("C211")
    ("cp 7")
    ("c 7")
    ("c 7")
    ("c 7")
    ("cp 7")
    ("c 6")
    ("c 7")
    ("cp 7")
  ))

(defparameter *DATOS-TURNO-D311*
  '(    ("D311" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("AE2" "c 2") ("RN" "c 2") ("TP 4" "") ("" "") ("" ""))
    ("Turno 2" ("MDE" "c 2") ("PL" "c 2") ("MDE" "cp 2") ("RN" "cp Lab2") (""
                                                                           ""))
    ("Turno 3" ("PGVD" "c 2") ("TP" "c 4") ("AE2" "cp Lab2") ("PGVD" "cp 7") (""
                                                                              ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-D311*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("AE2" "Análisis Estadístico II" "2" "0" "")
    ("MDE" "Muestreo y Diseño de Experimentos" "2" "0" "")
    ("RN" "Redes Neuronales" "2" "0" "")
    ("PL" "Procesamiento del Lenguaje" "2" "0" "")
    ("PGVD" "Procesamiento de Grandes Volúmenes de Datos" "2" "0" "")
    ("TP" "Teoría Política" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-D311*
  '(    ("")
    ("")
    ("Aulas")
    ("c 2")
    ("c 2")
    ("c 2")
    ("c 2")
    ("cp 2")
    ("cp Lab2")
    ("c 2")
    ("c 4")
    ("cp Lab2")
    ("cp 7")
  ))

(defparameter *DATOS-TURNO-D411*
  '(    ("D411" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "CP 2") ("EIA" "2") ("" "CP 2") ("SN" "C4") ("" ""))
    ("Turno 2" ("IN" "2") ("SN" "C4") ("IN" "2") ("SN" "C4") ("" ""))
    ("Turno 3" ("ECTS 9 (con" "C4") ("CO" "2") ("EIA 2" "") ("" "") ("" ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-D411*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("IN" "Inteligencia de Negocios" "2" "0" "")
    ("EIA" "Elementos de Inteligencia Artificial" "2" "0" "")
    ("CP" "Ciberseguridad y Privacidad" "2" "0" "")
    ("CO2" "Curso Optativo II" "2" "0" "")
    ("ECTS" "Estudios de Ciencia, Tecnología y Sociedad" "2" "0" "")
    ("SN/DN" "Seguridad Nacional / Defensa Nacional" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-D411*
  '(    ("")
    ("")
    ("Aulas")
    ("CP 2")
    ("2")
    ("CP 2")
    ("C4")
    ("2")
    ("C4")
    ("2")
    ("C4")
    ("C4")
    ("2")
  ))

(defparameter *DATOS-TURNO-C111*
  '(    ("C111" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("F" "Aula 6") ("A I" "Aula 6*") ("AM I" "Aula 6*") ("F"
                                                                    "Aula 6") (""
                                                                               ""))
    ("Turno 2" ("L" "Aula 6*") ("A I" "Aula 6*") ("EF" "SEDER") ("L" "Aula 6*") ("A I"
                                                                                 "Aula 6*"))
    ("Turno 3" ("P" "Aula 6") ("AM I" "Aula 6*") ("AM I" "Aula 6*") ("P"
                                                                     "Aula 6") ("P"
                                                                                "Lab"))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C111*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("A I" "Álgebra I" "3" "0" "")
    ("L" "Lógica" "2" "0" "")
    ("P" "Programación" "3" "0" "")
    ("AM I" "Análisis Matemático I" "2" "0" "")
    ("F" "Filosofía" "2" "0" "")
    ("EF" "Educación Física I" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C111*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 6")
    ("Aula 6*")
    ("Aula 6*")
    ("Aula 6")
    ("Aula 6*")
    ("Aula 6*")
    ("SEDER")
    ("Aula 6*")
    ("Aula 6*")
    ("Aula 6")
    ("Aula 6*")
    ("Aula 6*")
    ("Aula 6")
    ("Lab")
  ))

(defparameter *DATOS-TURNO-C121*
  '(    ("C121" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("P" "Aula 5") ("A I" "Aula 5") ("AM I" "Aula 5") ("P" "Aula 5") ("A I"
                                                                                 "Aula 5"))
    ("Turno 2" ("F" "Aula 5") ("AM I" "Aula 5") ("EF" "SEDER") ("L" "Aula 5") ("P"
                                                                               "Lab"))
    ("Turno 3" ("L" "Aula 5") ("A I" "Aula 5") ("AM I" "Aula 5") ("F" "Aula 5") (""
                                                                                 ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C121*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("A I" "Álgebra I" "3" "0" "")
    ("L" "Lógica" "2" "0" "")
    ("P" "Programación" "3" "0" "")
    ("AM I" "Análisis Matemático I" "2" "0" "")
    ("F" "Filosofía" "2" "0" "")
    ("EF" "Educación Física I" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C121*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
    ("SEDER")
    ("Aula 5")
    ("Lab")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
    ("Aula 5")
  ))

(defparameter *DATOS-TURNO-C122*
  '(    ("C122" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("P" "Aula 5") ("A I" "Aula 5") ("AM I" "Aula 1") ("L" "Aula 1") ("P"
                                                                                 "Lab"))
    ("Turno 2" ("F" "Aula 5") ("AM I" "Aula 5") ("EF" "SEDER") ("AM I" "Aula 1") ("A I"
                                                                                  "Aula 1"))
    ("Turno 3" ("L" "Aula 5") ("A I" "Aula 1") ("P" "Aula 1") ("F" "Aula 5") (""
                                                                              ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C122*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("A I" "Álgebra I" "3" "0" "")
    ("L" "Lógica" "2" "0" "")
    ("P" "Programación" "3" "0" "")
    ("AM I" "Análisis Matemático I" "2" "0" "")
    ("F" "Filosofía" "2" "0" "")
    ("EF" "Educación Física I" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C122*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 5")
    ("Aula 5")
    ("Aula 1")
    ("Aula 1")
    ("Lab")
    ("Aula 5")
    ("Aula 5")
    ("SEDER")
    ("Aula 1")
    ("Aula 1")
    ("Aula 5")
    ("Aula 1")
    ("Aula 1")
    ("Aula 5")
  ))

(defparameter *DATOS-TURNO-C211*
  '(    ("C211" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 4" ("EDA I" "c 6") ("MD" "c6") ("EDO" "cp 6") ("MN" "cp 6") ("AC"
                                                                         "lab"))
    ("Turno 5" ("TP" "c 6") ("MN" "c 6") ("EDA I" "cp 6") ("TP" "c 6") ("MD"
                                                                        "cp 6"))
    ("Turno 6" ("EDO" "c 6") ("AC" "c 6") ("EF 4:45pm a 5:35pm" "") ("" "") (""
                                                                             ""))
  ))

(defparameter *DATOS-STATS-C211*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("EDA" "Estructuras de Datos y Algoritmos I" "2" "0" "")
    ("MD" "Matemática Discreta I" "2" "0" "")
    ("AC" "Arquitectura de computadoras" "2" "0" "")
    ("EDO" "Ecuaciones Diferenciales Ordinarias" "2" "0" "")
    ("MN" "Matemática Numérica" "2" "0" "")
    ("TP" "Teoría Política" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C211*
  '(    ("")
    ("")
    ("Aulas")
    ("c 6")
    ("c6")
    ("cp 6")
    ("cp 6")
    ("lab")
    ("c 6")
    ("c 6")
    ("cp 6")
    ("c 6")
    ("cp 6")
    ("c 6")
    ("c 6")
  ))

(defparameter *DATOS-TURNO-C212*
  '(    ("C212" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 4" ("EDA I" "c 6") ("MD" "c6") ("EDO I" "cp 5") ("MN" "cp 5") ("AC"
                                                                           "lab"))
    ("Turno 5" ("TP" "c 6") ("MN" "c 6") ("EDA I" "cp 5") ("TP" "c 6") ("MD"
                                                                        "cp 5"))
    ("Turno 6" ("EDO" "c 6") ("AC" "c 6") ("EF 4:45pm a 5:35pm" "") ("" "") (""
                                                                             ""))
  ))

(defparameter *DATOS-STATS-C212*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("EDA" "Estructuras de Datos y Algoritmos I" "2" "0" "")
    ("MD" "Matemática Discreta I" "2" "0" "")
    ("AC" "Arquitectura de computadoras" "2" "0" "")
    ("EDO" "Ecuaciones Diferenciales Ordinarias" "2" "0" "")
    ("MN" "Matemática Numérica" "2" "0" "")
    ("TP" "Teoría Política" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C212*
  '(    ("")
    ("")
    ("Aulas")
    ("c 6")
    ("c6")
    ("cp 5")
    ("cp 5")
    ("lab")
    ("c 6")
    ("c 6")
    ("cp 5")
    ("c 6")
    ("cp 5")
    ("c 6")
    ("c 6")
  ))

(defparameter *DATOS-TURNO-C311*
  '(    ("C311" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("BD2" "Aula 9") ("Est" "Aula 9") ("BD2 cp" "Aula 9") ("MO"
                                                                      "Aula 9") (""
                                                                                 ""))
    ("Turno 2" ("IS c" "Aula 9") ("PD c" "Aula 9") ("IS c" "Aula 9") ("PD cp"
                                                                      "Aula 9") (""
                                                                                 ""))
    ("Turno 3" ("RC" "Aula 9") ("MO" "Aula 9") ("RC cp" "Aula 9") ("Est cp"
                                                                   "Aula 9") (""
                                                                              ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C311*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("RC" "Redes de Computadoras" "2" "0" "")
    ("IS" "Ingeniería de Software" "2" "0" "")
    ("MO" "Modelos de Optimización" "2" "0" "")
    ("BD2" "Bases de Datos II" "2" "0" "")
    ("PD" "Programación Declarativa" "2" "0" "")
    ("Est" "Estadística" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C311*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
    ("Aula 9")
  ))

(defparameter *DATOS-TURNO-C312*
  '(    ("C312" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("BD2" "Aula 9") ("Est" "Aula 9") ("BD2 cp" "Aula 3") ("PD cp"
                                                                      "Aula 3") (""
                                                                                 ""))
    ("Turno 2" ("IS c" "Aula 9") ("PD c" "Aula 9") ("IS c" "Aula 3") ("MO"
                                                                      "Aula 3") (""
                                                                                 ""))
    ("Turno 3" ("RC" "Aula 9") ("MO" "Aula 9") ("RC cp" "Aula 3") ("Est cp"
                                                                   "Aula 3") (""
                                                                              ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C312*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("RC" "Redes de Computadoras" "2" "0" "")
    ("IS" "Ingeniería de Software" "2" "0" "")
    ("MO" "Modelos de Optimización" "2" "0" "")
    ("BD2" "Bases de Datos II" "2" "0" "")
    ("PD" "Programación Declarativa" "2" "0" "")
    ("Est" "Estadística" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C312*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 9")
    ("Aula 9")
    ("Aula 3")
    ("Aula 3")
    ("Aula 9")
    ("Aula 9")
    ("Aula 3")
    ("Aula 3")
    ("Aula 9")
    ("Aula 9")
    ("Aula 3")
    ("Aula 3")
  ))

(defparameter *DATOS-TURNO-C411*
  '(    ("C411" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 4" ("AM" "9") ("DAA" "9") ("DAA" "9") ("SN" "") ("" ""))
    ("Turno 5" ("SD 9" "") ("SN" "") ("AM" "9") ("SN" "") ("" ""))
    ("Turno 6" ("ECTS 9" "") ("AE" "9") ("SD" "9") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C411*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("AM" "Aprendizaje de Máquinas" "2" "0" "")
    ("DAA" "Diseño y Análisis de Algoritmos" "2" "0" "")
    ("SD" "Sistemas Distribuidos" "2" "0" "")
    ("AE" "Asignatura Electiva" "2" "0" "")
    ("ECTS" "Estudios de Ciencia, Tecnología y Sociedad" "2" "0" "")
    ("SN/DN" "Seguridad Nacional / Defensa Nacionaol" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C411*
  '(    ("")
    ("")
    ("Aulas")
    ("9")
    ("9")
    ("9")
    ("9")
    ("9")
    ("9")
  ))

(defparameter *DATOS-TURNO-C412*
  '(    ("C412" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 4" ("AM" "9") ("DAA" "9") ("DAA" "9") ("SN" "") ("" ""))
    ("Turno 5" ("SD 9" "") ("SN" "") ("AM" "1") ("SN" "") ("" ""))
    ("Turno 6" ("ECTS 9" "") ("AE" "") ("SD" "9") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-C412*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("AM" "Aprendizaje de Máquinas" "2" "0" "")
    ("DAA" "Diseño y Análisis de Algoritmos" "2" "0" "")
    ("SD" "Sistemas Distribuidos" "2" "0" "")
    ("AE" "Asignatura Electiva" "2" "0" "")
    ("ECTS" "Estudios de Ciencia, Tecnología y Sociedad" "2" "0" "")
    ("SN/DN" "Seguridad Nacional / Defensa Nacionaol" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-C412*
  '(    ("")
    ("")
    ("Aulas")
    ("9")
    ("9")
    ("9")
    ("1")
    ("9")
  ))

(defparameter *DATOS-TURNO-M111*
  '(    ("M111" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("F" "Aula 8") ("IM" "Aula 8") ("IAM" "Aula 8") ("IA" "Aula 8") ("IA"
                                                                                "Aula 8"))
    ("Turno 2" ("PA" "Aula 8") ("IAM" "Aula 8") ("EF" "SEDER") ("PA Aula" "Lab") ("IAM"
                                                                                  "Aula 8"))
    ("Turno 3" ("GA" "Aula 8") ("GA" "Aula 8") ("IA" "Aula 8") ("GA 8" "") (""
                                                                            ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-M111*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("IAM" "Introducción al Análisis Matemático" "3" "0" "")
    ("IA" "Introducción al Álgebra" "3" "0" "")
    ("GA" "Geometría Analítica" "2" "0" "")
    ("PA" "Programación y Algoritmos" "2" "0" "")
    ("IM" "Introducción a la Matemática" "2" "0" "")
    ("F" "Filosofía" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-M111*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
    ("SEDER")
    ("Lab")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
    ("Aula 8")
  ))

(defparameter *DATOS-TURNO-M211*
  '(    ("M211" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("CAL" "c 3") ("CAL" "c 3") ("" "") ("" "") ("" ""))
    ("Turno 4" ("FVV" "c 3") ("FVV" "c 3") ("FVV" "cp 3") ("CAL" "cp 3") ("" ""))
    ("Turno 5" ("EP" "c 7") ("EP" "c 7") ("FVV" "cp 3") ("" "") ("" ""))
    ("Turno 6" ("EF 4:45pm a 5:35pm" "") ("SP" "2") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-M211*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("FVV" "Funciones de Varias Variables" "4" "0" "")
    ("CAL" "Complementos de Álgebra Lineal" "3" "0" "")
    ("SP2" "Seminario de Problemas II" "2" "0" "")
    ("AE" "Asignatura Electiva I" "2" "0" "")
    ("EP" "Economía Política" "2" "0" "")
    ("EF" "Educación Física III" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-M211*
  '(    ("")
    ("")
    ("Aulas")
    ("c 3")
    ("c 3")
    ("c 3")
    ("c 3")
    ("cp 3")
    ("cp 3")
    ("c 7")
    ("c 7")
    ("cp 3")
    ("2")
  ))

(defparameter *DATOS-TURNO-M311*
  '(    ("M311" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("MN" "Aula 4") ("EDO" "Aula 4") ("MN" "4") ("FV" "C 4") ("FVC"
                                                                         "Aula 4"))
    ("Turno 2" ("FVC" "Aula 4") ("OM" "Aula 4") ("EDO" "Aula 4") ("IE" "Aula 4") ("EDO"
                                                                                  "Aula 4"))
    ("Turno 3" ("IE" "Aula 4") ("TP" "Aula 4") ("IE" "Aula 4") ("OM" "Aula 4") (""
                                                                                ""))
    ("Turno 4" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 5" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 6" ("" "") ("" "") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-M311*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("FVC" "Funciones de variable Compleja" "2" "0" "")
    ("IE" "Inferencia Estadística" "3" "0" "")
    ("EDO" "Ecuaciones Diferenciales Ordinarias" "3" "0" "")
    ("MN" "Matemática Numérica" "2" "0" "")
    ("OM" "Optimización Matemática I" "2" "0" "")
    ("TP" "Teoría Política" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-M311*
  '(    ("")
    ("")
    ("Aulas")
    ("Aula 4")
    ("Aula 4")
    ("4")
    ("C 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
    ("Aula 4")
  ))

(defparameter *DATOS-TURNO-M411*
  '(    ("M411" "" "" "" "" "")
    ("" "" "" "" "" "")
    ("" "Lunes" "Martes" "Miercoles" "Jueves" "Viernes")
    ("Turno 1" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 2" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 3" ("" "") ("" "") ("" "") ("" "") ("" ""))
    ("Turno 4" ("GD" "4") ("HM" "4") ("MI" "4") ("AO" "2") ("" ""))
    ("Turno 5" ("MI" "4") ("HM" "4") ("GD" "4") ("AO" "3") ("" ""))
    ("Turno 6" ("ECTS 9 (Con" "C4") ("GD" "4") ("" "") ("" "") ("" ""))
  ))

(defparameter *DATOS-STATS-M411*
  '(    ("" "" "" "" "")
    ("" "" "" "" "")
    ("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas")
    ("MI" "Medida e Integración" "2" "0" "")
    ("GD" "Geometría Diferencial" "2" "0" "")
    ("HM" "Historia de la Matemática" "2" "0" "")
    ("ECTS" "Estudios de Ciencia, Tecnología y Sociedad" "2" "0" "")
    ("AO2" "Asignatura Optativa II" "2" "0" "")
    ("AO3" "Asignatura Optativa III" "2" "0" "")
  ))

(defparameter *DATOS-AULAS-M411*
  '(    ("")
    ("")
    ("Aulas")
    ("4")
    ("4")
    ("4")
    ("2")
    ("4")
    ("4")
    ("4")
    ("3")
    ("C4")
    ("4")
  ))

(defparameter *DATOS-AULAS-LUNES*
  '(    ("Lunes" "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
    ("1" "" "" "" "" "" "C111,C112" "" "" "" "C113")
    ("2" "" "" "" "" "C112" "C113" "" "" "" "")
    ("3" "" "" "" "" "" "C111,C113" "" "C112" "" "")
    ("4" "" "" "" "" "" "" "" "" "" "")
    ("5" "" "" "" "" "" "" "" "" "" "")
    ("6" "" "" "" "" "" "" "" "" "" "")
  ))

(defparameter *DATOS-AULAS-MARTES*
  '(    ("Martes" "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
    ("1" "" "C112" "" "" "" "C111" "C113" "" "" "")
    ("2" "C111" "" "" "" "" "C112,C113" "" "" "" "")
    ("3" "" "" "" "" "" "C111,C112,C113" "" "" "" "")
    ("4" "" "" "" "" "" "" "" "" "" "")
    ("5" "" "" "" "" "" "" "" "" "" "")
    ("6" "" "" "" "" "" "" "" "" "" "")
  ))

(defparameter *DATOS-AULAS-MIERCOLES*
  '(    ("Miercoles" "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
    ("1" "" "" "" "" "C112" "C111" "" "C113" "" "")
    ("2" "" "" "" "" "" "" "" "" "" "")
    ("3" "" "C112" "" "" "C113" "C111" "" "" "" "")
    ("4" "" "" "" "" "" "" "" "" "" "")
    ("5" "" "" "" "" "" "" "" "" "" "")
    ("6" "" "" "" "" "" "" "" "" "" "")
  ))

(defparameter *DATOS-AULAS-JUEVES*
  '(    ("Jueves" "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
    ("1" "C113" "" "" "" "C112" "C111" "" "" "" "")
    ("2" "C113" "" "" "" "C112" "C111" "" "" "" "")
    ("3" "C113" "" "" "" "C112" "C111" "" "" "" "")
    ("4" "" "" "" "" "" "" "" "" "" "")
    ("5" "" "" "" "" "" "" "" "" "" "")
    ("6" "" "" "" "" "" "" "" "" "" "")
  ))

(defparameter *DATOS-AULAS-VIERNES*
  '(    ("Viernes" "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
    ("1" "" "" "" "" "" "" "" "" "" "")
    ("2" "" "" "" "" "" "" "" "" "" "C112")
    ("3" "" "" "" "" "C111" "" "" "" "" "")
    ("4" "" "" "" "" "" "" "" "" "" "")
    ("5" "" "" "" "" "" "" "" "" "" "")
    ("6" "" "" "" "" "" "" "" "" "" "")
  ))

