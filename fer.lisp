(def-global-data datos-generales
    ;; aquí definimos 
    ;; aquí solo tenemos como dato global la lista de programa
    (tabla tipos-de-programas
           (id nombre))
  (tabla programas 
         (id nombre duracion tipo-de-publico (tipo tipos-de-programas))))



;; esta es la tabla que muestra el horario de un día
(def-tabla horario-diario (dia hora-de-inicio-del-primer-programa)
  :name dia
  :columns ((hora-inicio "Hora Inicio")
            (hora-final "Hora Final")
            (nombre-del-programa "Nombre"))
  ;; la altura de cada «celda»
  :height 1
  :local-data hora-de-inicio-del-primer-programa
  ;; esto que sigue hay que pensarlo un poquito
  ;; pero justamente por eso se llama investigación 😅
  ;; ahora mismo (en mi cabeza) column-compute define una función
  ;; que la aplica la fórmula a toda la columna
  ;; el primer argumento es el nombre de la columna
  ;; el resto de los arg pueden ser opcionales
  ;; y devuelve qué calcular ahí
  :formula (column-compute hora-final (row)
            ;; indica cómo calcular la hora final de una fila                
                    ;; no puede ser if porque if está reservado 😒
                    ;; pero es una condicional
                    (if (non-empty (nombre-del-programa row))
                        (then
                         (time-add (hora-incio row)
                                        (duracion nombre-del-programa row)))
                        (else
                         show-nothing)))
  ;; tampoco se puede tener dos fórmulas así,
  ;; pero no quiero meterlo dentro de una lista ahora
  :formula (column-compute hora-inicial (r)
            ;; aquí definimos cómo calcular la hora inicial               
                           (if (it-is-the-first-row r)
                               ;; si es la primera fila
                               (then
                                ;; devuelve el dato que declaré por ahí.
                                hora-de-inicio-del-primer-programa)
                               ;; en otro caso
                               (else
                                ;; si la fila anterior tiene nombre del programa
                                (if non-empty (nombre-del-programa (previous-row r))
                                    (then
                                     ;; entonces la hora inicial es la final de la fila anterior
                                     (hora-final (previous-row row))
                                     (else ;; esto es que la fila anterior no tiene programa
                                      ;; si no, déjalo en blanco
                                      show-nothing))))))
  ;; ahora un poquito de formato condicional
  :cond-format (column-format nombre-del-programa (r)
      ;; este primero es para resaltar en rojo que alguien
      ;; ponga un programa que no existe                        
                  (if (is-not-defined (nombre-del-programa r) :in (nombre tabla-programas))
                      (then
                        (set-background-color r red))))
  ;; y otro más
  :cond-format (column-format nombre-del-programa (r)
                              ;; esto significa que no lo apliques a la primera fila
                              :apply-from (column 2)
      ;; este es para que no haya dos del mismo tipo de programa seguidos
                              (if (is-the-same
                                    (tipo (nombre-del-programa (previous-row r)))
                                    (tipo (nombre-del-programa r)))
                                  (then
                                   (set-background-color r orange)
                                   (set-background-color (previous-row r) orange)))))

(def-hoja planificacion-diaria (nombre)
  ;; aquí vamos a definir cómo se vería una hoja
  ;; una hoja tiene un nombre
  :dia nombre
  ;; tiene un dato que se va a llamar hora-de-inicio-del-dia
  ;; y que es de tipo time
  :data (hora-de-inicio-del-dia :de-tipo time) Fernando Rodríguez, [May 13, 2026 at 7:26 PM]


  ;; y lo otro que tiene es una tabla de tipo horario-diario
  ;; que recibe como argumento el nombre del día
  ;; y la hora de inicio del día
  (horario-diario nombre hora-de-inicio-del-dia))

(def-horario planificacion-TV
    :hojas (list
            (planificacion-diaria "Lunes")
            (planificacion-diaria "Martes")
            (planificacion-diaria "Miércoles")
            (planificacion-diaria "Jueves")
            (planificacion-diaria "Viernes")
            (planificacion-diaria "Sábado")
            (planificacion-diaria "Domingo")

            ;; esto significa que el comando
            ;; def-global-data, además de generar los datos globales
            ;; genera una hoja con ese nombre
            datos-generales))