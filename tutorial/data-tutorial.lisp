;; data-tutorial.lisp
;; Datos reales del horario de defensas de tesis — junio 2026.
;;
;; Fuente: "Versión final Tribunales (junio) 25-26.xlsx"
;;
;; Columnas defensas: dia | hora | estudiante | tutor | presidente |
;;                    secretario | vocal | oponente | local
;;
;; Columnas profesores: nombre | defensas (calculada)
;;
;; Nota: cuando un estudiante tenía varios tutores se conservó únicamente
;;       el primero de la celda original.

;; =====================================================================
;; TABLA DE DEFENSAS — una fila por defensa
;; =====================================================================

(defparameter *DATOS-DEFENSAS*
  '(("" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "")
    ("Día" "Hora" "Estudiante" "Tutor"
     "Presidente" "Secretario" "Vocal" "Oponente" "Local")
    ;; --- 5 de junio ---
    ("05/06/2026" "10:00"
     "Angel Daniel Alonso Guevara"
     "MSc. Carmen T. Fernández Montoto"
     "Dr. Alberto Fernández Oliva"
     "Lic. Alejandro Labourdette-Lantigua Soto"
     "Lic. Dennis Daniel González Durán"
     "MSc. Aracelys García Armenteros"
     "Postgrado")
    ;; --- 8 de junio ---
    ("08/06/2026" "09:30"
     "Adrián Hernández Castellanos"
     "Lic. Alejandra Monzón Peña"
     "Lic. Amanda Noris Hernández"
     "Lic. Kevin Manzano Rodríguez"
     "Lic. Daniel Abad Fundora"
     "Lic. Rodrigo García Gómez"
     "Salón decanato")
    ("08/06/2026" "10:30"
     "Laura Martir Beltrán"
     "Lic. Alejandra Monzón Peña"
     "Lic. Amanda Noris Hernández"
     "Lic. Kevin Manzano Rodríguez"
     "Lic. Daniel Abad Fundora"
     "Lic. Rodrigo García Gómez"
     "Salón decanato")
    ("08/06/2026" "11:30"
     "Noel Pérez Calvo"
     "MSc. Fernando Raúl Rodríguez Flores"
     "Lic. Rodrigo García Gómez"
     "Lic. Lázaro Daniel González Martínez"
     "Lic. Alejandra Monzón Peña"
     "Lic. David Guaty Domínguez"
     "Salón decanato")
    ;; --- 10 de junio (Postgrado) ---
    ("10/06/2026" "09:30"
     "Claudia Hernández Pérez y Joel Aparicio Tamayo"
     "MSc. Celia T. González González"
     "Dra. Ayme Marrero Severo"
     "MSc. Joanna Campbell Amos"
     "Lic. Daniel Alejandro Valdés Pérez"
     "Lic. Alejandro Beltrán Varela"
     "Postgrado")
    ("10/06/2026" "11:00"
     "Darío Hernández Cubilla"
     "MSc. Aracelys García Armenteros"
     "Dra. Ayme Marrero Severo"
     "Lic. Roberto Marti Cedeño"
     "Lic. Alejandro Beltrán Varela"
     "Lic. Alejandro Labourdette-Lantigua Soto"
     "Postgrado")
    ("10/06/2026" "12:00"
     "Richard Alejandro Matos Arderí"
     "MSc. Celia T. González González"
     "MSc. Wilfredo Morales Lezca"
     "Lic. Roberto Marti Cedeño"
     "Lic. Alejandro Beltrán Varela"
     "MSc. Aracelys García Armenteros"
     "Postgrado")
    ;; --- 10 de junio (Francofonía) ---
    ("10/06/2026" "09:30"
     "Melani Forsythe Matos"
     "DraC. Marta Lourdes Baguer Díaz-Romañach"
     "DrC. Alejandro Sánchez Castellanos"
     "Lic. Ana Paula González Muñoz"
     "Lic. Dennis Daniel González Durán"
     "Lic. Javier Rodríguez Sánchez"
     "Francofonía")
    ("10/06/2026" "10:30"
     "Mauro Eduardo Campver Barrios"
     "DraC. Marta Lourdes Baguer Díaz-Romañach"
     "DrC. Alejandro Piad Morfis"
     "Lic. Ana Paula González Muñoz"
     "Lic. Dennis Daniel González Durán"
     "MSc. Fernando Raúl Rodríguez Flores"
     "Francofonía")
    ("10/06/2026" "13:30"
     "José Miguel Leyva De La Cruz"
     "Lic. Alejandro Beltrán Varela"
     "Lic. Amanda Noris Hernández"
     "Lic. Juan Miguel Pérez Martínez"
     "Dr. Gabriel Fundora"
     "MSc. Aracelys García Armenteros"
     "Francofonía")
    ;; --- 10 de junio (Salón decanato) ---
    ("10/06/2026" "11:30"
     "Jossué Arteche Muñoz"
     "MSc. Fernando Raúl Rodríguez Flores"
     "DrC. Alejandro Piad Morfis"
     "Lic. Rodrigo García Gómez"
     "Lic. Daniel Abad Fundora"
     "Lic. Ariel González Gómez"
     "Salón decanato")
    ("10/06/2026" "12:30"
     "Amalia Beatriz Valiente Hinojosa"
     "MSc. Fernando Raúl Rodríguez Flores"
     "Lic. Javier Rodríguez Sánchez"
     "Lic. Rocío Ortiz Gancedo"
     "Lic. Merling Sabater Ramírez"
     "Lic. Alejandra Monzón Peña"
     "Salón decanato")
    ("10/06/2026" "13:30"
     "Luis Ernesto Amat Cárdenas"
     "Lic. Roberto Marti Cedeño"
     "DrC. Alejandro Piad Morfis"
     "Lic. Carlos David Muñiz Chall"
     "Lic. Alejandro Labourdette-Lantigua Soto"
     "Lic. Daniel Toledo Martínez"
     "Salón decanato")))

;; =====================================================================
;; TABLA DE PROFESORES — una fila por profesor
;; El nombre incluye el grado académico para coincidir exactamente con
;; los valores almacenados en las columnas de tribunal de *DATOS-DEFENSAS*.
;; La columna "defensas" se calcula con COUNTIF en tiempo de compilación.
;; =====================================================================

(defparameter *DATOS-PROFESORES*
  '(("" "")
    ("" "")
    ("Nombre" "Defensas")
    ("Dr. Alberto Fernández Oliva"                "")
    ("Dr. Gabriel Fundora"                        "")
    ("DrC. Alejandro Piad Morfis"                 "")
    ("DrC. Alejandro Sánchez Castellanos"         "")
    ("Dra. Ayme Marrero Severo"                   "")
    ("DraC. Marta Lourdes Baguer Díaz-Romañach"   "")
    ("Lic. Alejandra Monzón Peña"                 "")
    ("Lic. Alejandro Beltrán Varela"              "")
    ("Lic. Alejandro Labourdette-Lantigua Soto"   "")
    ("Lic. Amanda Noris Hernández"                "")
    ("Lic. Ana Paula González Muñoz"              "")
    ("Lic. Ariel González Gómez"                  "")
    ("Lic. Carlos David Muñiz Chall"              "")
    ("Lic. Daniel Abad Fundora"                   "")
    ("Lic. Daniel Alejandro Valdés Pérez"         "")
    ("Lic. Daniel Toledo Martínez"                "")
    ("Lic. David Guaty Domínguez"                 "")
    ("Lic. Dennis Daniel González Durán"          "")
    ("Lic. Javier Rodríguez Sánchez"              "")
    ("Lic. Juan Miguel Pérez Martínez"            "")
    ("Lic. Kevin Manzano Rodríguez"               "")
    ("Lic. Lázaro Daniel González Martínez"       "")
    ("Lic. Merling Sabater Ramírez"               "")
    ("Lic. Roberto Marti Cedeño"                  "")
    ("Lic. Rocío Ortiz Gancedo"                   "")
    ("Lic. Rodrigo García Gómez"                  "")
    ("MSc. Aracelys García Armenteros"            "")
    ("MSc. Carmen T. Fernández Montoto"           "")
    ("MSc. Celia T. González González"            "")
    ("MSc. Fernando Raúl Rodríguez Flores"        "")
    ("MSc. Joanna Campbell Amos"                  "")
    ("MSc. Wilfredo Morales Lezca"                "")))
