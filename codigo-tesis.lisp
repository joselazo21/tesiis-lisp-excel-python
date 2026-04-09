(defun mkstr (&rest args)
     "Returns a string with the concatenation of the args"
     (string-upcase
      (with-output-to-string (s)
        (dolist (a args) (princ a s)))))

(defun symb (&rest args)
     "Returns a symbol formed by the concatenation of the args."
     (values (intern (apply #'mkstr args))))

(defun keywd (&rest args)
     "Returns a keyword formed by the concatenation of the args."
     (values (intern (apply #'mkstr args) :keyword)))

(defun ppme (expr &optional (stream t))
  (format stream "~s" (macroexpand-1 expr)))

(defun expansion-del-slot (s)
  `(,s :accessor ,s :initarg ,(keywd s)))

(defun make-constructor (class-name slots &key (argument-type nil))
  `(defun ,class-name ,(append (if argument-type 
       			    (list argument-type))
       		 slots)
     (make-instance ',(symb 'clase- class-name)
         ,@(loop for s in slots
              collect (keywd s)
              collect s))))

(defmacro defclass* (class-name parents slots
                      &key (ctr-args-type '&key))
  `(progn
     (defclass ,(symb 'clase- class-name) ,parents
       ,(loop for s in slots
              collecting (expansion-del-slot s)))
     ,(make-constructor class-name slots :argument-type ctr-args-type)))

(defclass* tabla ()
    (id
     filas
     columnas
     alto-celda
     ancho-celda
     nombres-columnas
     nombres-filas
  contenido-de-la-tabla
  border-color
  border-style
  range-styles))

  (defmethod print-object ((obj clase-tabla) stream)
    (if (contenido-de-la-tabla obj)
        (format stream "[Tabla (con contenido): ~a x ~a]"
            (length (contenido-de-la-tabla obj))
            (length (first (contenido-de-la-tabla obj))))
        ;; else
        (format stream "[Tabla: ~a x ~a]"
            (filas obj)
            (columnas obj))))

 ;; ponemos un initialize-instance :after
(defmethod initialize-instance :after ((obj clase-tabla) 
       				&key &allow-other-keys)
  
  (if (contenido-de-la-tabla obj)
      ;; si nos pasaron contenido-de-la-tabla, sobreescribe todo lo demás
      (progn   
         (setf (columnas obj) (length (first (contenido-de-la-tabla obj))))
         (setf (filas obj) (length (contenido-de-la-tabla obj))))
      ;; else
      (progn
        (if (nombres-columnas obj)
            ;; si nos pasaron los nombres de las columnas, sobreescribe columnas
            (setf (columnas obj) (length (nombres-columnas obj))))
        (if (nombres-filas obj)
            ;; si nos pasaron los nombres de las filas, sobreescribe columnas
            (setf (filas obj) (length (nombres-filas obj))))))
  (if (null (columnas obj))
      (error "Error al crear tabla.  Se debe especificar columnas, nombres-columnas, o contenido-de-la-tabla."))
  (if (null (filas obj))
      (error "Error al crear tabla.  Se debe especificar filas, nombres-filas, o contenido-de-la-tabla.")))

(defclass* hoja ()
  (grupo
   horario
   asignaturas))

(defmethod print-object ((obj clase-hoja) stream)
  (format stream "[Hoja: ~a, con ~a asignaturas]"
          (grupo obj)
          (filas (asignaturas obj))))

(defclass* libro ()
  (hojas))

(defmethod print-object ((obj clase-libro) stream)
  (format stream "[Libro con ~a hojas]"
          (length (hojas obj))))

(defun tabla-horarios-para-todo-el-mundo (grupo)
  (tabla :id (concatenate 'string "Horario-" grupo)
	  :nombres-columnas '("Turnos"
                              "Lunes"
                              "Martes"
                              "Miércoles"
                              "Jueves"
                              "Viernes")
	  :alto-celda 3
	  :nombres-filas '("Turno 1"
                           "Turno 2"
                           "Turno 3"
                           "Turno 4"
                           "Turno 5"
                           "Turno 6")))

(defparameter *columnas-fijas-horario-aula*
  '("Turnos" "Lunes" "Martes" "Miércoles" "Jueves" "Viernes"))

(defun nombres-filas-por-turnos (cantidad-turnos &optional (alto-celda 1))
  "Genera nombres de filas: Turno 1..n y subfilas vacias segun alto-celda."
  (let ((alto (max 1 alto-celda))
        (result '()))
    (loop for i from 1 to cantidad-turnos
          do (push (format nil "Turno ~a" i) result)
             (loop repeat (1- alto)
                   do (push "" result)))
    (nreverse result)))

(defmacro tabla-horario-aula (&key
                                id
                                cantidad-turnos
                                (alto-celda 1)
                                (ancho-celda 1)
                                contenido-de-la-tabla
                                (border-color "000000")
                                (border-style "medium")
                                (range-styles nil))
  "Construye una tabla de horario para un aula con columnas fijas y turnos parametrizables.

Recibe alto/ancho de celda, contenido, color/estilo de borde y estilos por rango.
range-styles usa la misma idea que en Python: una lista de configuraciones por rango."
  `(tabla :id ,id
          :alto-celda ,alto-celda
          :ancho-celda ,ancho-celda
          :columnas (length *columnas-fijas-horario-aula*)
          :nombres-columnas *columnas-fijas-horario-aula*
          :nombres-filas (nombres-filas-por-turnos ,cantidad-turnos ,alto-celda)
          :contenido-de-la-tabla ,contenido-de-la-tabla
          :border-color ,border-color
          :border-style ,border-style
          :range-styles ,range-styles))

(defmacro def-tabla-de-asignaturas (anno asignaturas)
  `(defun ,(symb 'tabla-asignaturas- anno) (grupo)
     ,(mkstr "Crea la tabla de asignaturas para " anno)
     (tabla :id (concatenate 'string "Asignaturas-" grupo)
         :contenido-de-la-tabla ',asignaturas)))

(def-tabla-de-asignaturas d1
  (("Álgebra Lineal-C" 1)
    ("Álgebra Lineal-CP" 1)
    ("Lógica-C" 1)
    ("Lógica-CP" 1)
    ("Introducción a la Programación-C" 1)
    ("Introducción a la Programación-CP" 1)
    ("Análisis Matemático I-C" 1)
    ("Análisis Matemático I-CP" 1)
    ("Introducción a la Ciencia de Datos-C" 1)
    ("Introducción a la Ciencia de Datos-CP" 1)
    ("Filosofía-C" 1)
    ("Filosofía-CP" 1)
    ("Educación Física I-C" 1)
    ("Educación Física I-CP" 1)))

(def-tabla-de-asignaturas d2
  (("Matemática y Aplicaciones-C" 1)
    ("Matemática y Aplicaciones-CP" 1)
    ("Probabilidades-C" 1)
    ("Probabilidades-CP" 1)
    ("Bases de Datos-C" 1)
    ("Bases de Datos-CP" 1)
    ("Estructura de Datos-C" 1)
    ("Estructura de Datos-CP" 1)
    ("Visualización de Datos-C" 1)
    ("Visualización de Datos-CP" 1)
    ("Economía Política-C" 1)
    ("Economía Política-CP" 1)
    ("Educación Física III-C" 1)
    ("Educación Física III-CP" 1)))

(def-tabla-de-asignaturas d3
  (("Análisis Estadístico II-C" 1)
    ("Análisis Estadístico II-CP" 1)
    ("Muestreo y Diseño de Experimentos-C" 1)
    ("Muestreo y Diseño de Experimentos-CP" 1)
    ("Redes Neuronales-C" 1)
    ("Redes Neuronales-CP" 1)
    ("Procesamiento del Lenguaje-C" 1)
    ("Procesamiento del Lenguaje-CP" 1)
    ("Procesamiento de Grandes Volúmenes de Datos-C" 1)
    ("Procesamiento de Grandes Volúmenes de Datos-CP" 1)
    ("Teoría Política-C" 1)
    ("Teoría Política-CP" 1)))

(def-tabla-de-asignaturas d4
  (("Inteligencia de Negocios-C" 1)
    ("Inteligencia de Negocios-CP" 1)
    ("Elementos de Inteligencia Artificial-C" 1)
    ("Elementos de Inteligencia Artificial-CP" 1)
    ("Ciberseguridad y Privacidad-C" 1)
    ("Ciberseguridad y Privacidad-CP" 1)
    ("Curso Optativo II-C" 1)
    ("Curso Optativo II-CP" 1)
    ("Estudios de Ciencia, Tecnología y Sociedad-C" 1)
    ("Estudios de Ciencia, Tecnología y Sociedad-CP" 1)
    ("Seguridad Nacional / Defensa Nacional-C" 1)
    ("Seguridad Nacional / Defensa Nacional-CP" 1)))

(def-tabla-de-asignaturas c1
  (("Álgebra I-C" 1)
    ("Álgebra I-CP" 1)
    ("Lógica-C" 1)
    ("Lógica-CP" 1)
    ("Programación-C" 1)
    ("Programación-CP" 1)
    ("Análisis Matemático I-C" 1)
    ("Análisis Matemático I-CP" 1)
    ("Filosofía-C" 1)
    ("Filosofía-CP" 1)
    ("Educación Física I-C" 1)
    ("Educación Física I-CP" 1)))

(def-tabla-de-asignaturas c2
  (("Estructuras de Datos y Algoritmos I-C" 1)
    ("Estructuras de Datos y Algoritmos I-CP" 1)
    ("Matemática Discreta I-C" 1)
    ("Matemática Discreta I-CP" 1)
    ("Arquitectura de Computadoras-C" 1)
    ("Arquitectura de Computadoras-CP" 1)
    ("Ecuaciones Diferenciales Ordinarias-C" 1)
    ("Ecuaciones Diferenciales Ordinarias-CP" 1)
    ("Matemática Numérica-C" 1)
    ("Matemática Numérica-CP" 1)
    ("Teoría Política-C" 1)
    ("Teoría Política-CP" 1)
    ("Educación Física III-C" 1)
    ("Educación Física III-CP" 1)))

(def-tabla-de-asignaturas c3
  (("Redes de Computadoras-C" 1)
    ("Redes de Computadoras-CP" 1)
    ("Ingeniería de Software-C" 1)
    ("Ingeniería de Software-CP" 1)
    ("Modelos de Optimización-C" 1)
    ("Modelos de Optimización-CP" 1)
    ("Bases de Datos II-C" 1)
    ("Bases de Datos II-CP" 1)
    ("Programación Declarativa-C" 1)
    ("Programación Declarativa-CP" 1)
    ("Estadística-C" 1)
    ("Estadística-CP" 1)))

(def-tabla-de-asignaturas c4
  (("Aprendizaje de Máquinas-C" 1)
    ("Aprendizaje de Máquinas-CP" 1)
    ("Diseño y Análisis de Algoritmos-C" 1)
    ("Diseño y Análisis de Algoritmos-CP" 1)
    ("Sistemas Distribuidos-C" 1)
    ("Sistemas Distribuidos-CP" 1)
    ("Asignatura Electiva-C" 1)
    ("Asignatura Electiva-CP" 1)
    ("Estudios de Ciencia, Tecnología y Sociedad-C" 1)
    ("Estudios de Ciencia, Tecnología y Sociedad-CP" 1)
    ("Seguridad Nacional / Defensa Nacional-C" 1)
    ("Seguridad Nacional / Defensa Nacional-CP" 1)))

;; MATEMÁTICA

(def-tabla-de-asignaturas m1
  (("Introducción al Análisis Matemático-C" 1)
    ("Introducción al Análisis Matemático-CP" 1)
    ("Introducción al Álgebra-C" 1)
    ("Introducción al Álgebra-CP" 1)
    ("Geometría Analítica-C" 1)
    ("Geometría Analítica-CP" 1)
    ("Programación y Algoritmos-C" 1)
    ("Programación y Algoritmos-CP" 1)
    ("Introducción a la Matemática-C" 1)
    ("Introducción a la Matemática-CP" 1)
    ("Filosofía-C" 1)
    ("Filosofía-CP" 1)
    ("Educación Física I-C" 1)
    ("Educación Física I-CP" 1)))

(def-tabla-de-asignaturas m2
  (("Funciones de Varias Variables-C" 1)
    ("Funciones de Varias Variables-CP" 1)
    ("Complementos de Álgebra Lineal-C" 1)
    ("Complementos de Álgebra Lineal-CP" 1)
    ("Seminario de Problemas II-C" 1)
    ("Seminario de Problemas II-CP" 1)
    ("Asignatura Electiva I-C" 1)
    ("Asignatura Electiva I-CP" 1)
    ("Economía Política-C" 1)
    ("Economía Política-CP" 1)
    ("Educación Física III-C" 1)
    ("Educación Física III-CP" 1)))

(def-tabla-de-asignaturas m3
  (("Funciones de Variable Compleja-C" 1)
    ("Funciones de Variable Compleja-CP" 1)
    ("Inferencia Estadística-C" 1)
    ("Inferencia Estadística-CP" 1)
    ("Ecuaciones Diferenciales Ordinarias-C" 1)
    ("Ecuaciones Diferenciales Ordinarias-CP" 1)
    ("Matemática Numérica-C" 1)
    ("Matemática Numérica-CP" 1)
    ("Optimización Matemática I-C" 1)
    ("Optimización Matemática I-CP" 1)
    ("Teoría Política-C" 1)
    ("Teoría Política-CP" 1)))

(def-tabla-de-asignaturas m4
  (("Medida e Integración-C" 1)
    ("Medida e Integración-CP" 1)
    ("Geometría Diferencial-C" 1)
    ("Geometría Diferencial-CP" 1)
    ("Historia de la Matemática-C" 1)
    ("Historia de la Matemática-CP" 1)
    ("Estudios de Ciencia, Tecnología y Sociedad-C" 1)
    ("Estudios de Ciencia, Tecnología y Sociedad-CP" 1)
    ("Asignatura Optativa II-C" 1)
    ("Asignatura Optativa II-CP" 1)
    ("Asignatura Optativa III-C" 1)
    ("Asignatura Optativa III-CP" 1)))

(defmacro def-hoja (anno)
  `(defun ,(symb 'hoja- anno) (grupo)
    (let* ((nombre-del-grupo (mkstr ,(mkstr anno) grupo)))
     (hoja :grupo nombre-del-grupo
         :horario (tabla-horarios-para-todo-el-mundo nombre-del-grupo)
         :asignaturas (,(symb 'tabla-asignaturas- anno) nombre-del-grupo)))))

(progn
  (def-hoja c1)
  (def-hoja c2)
  (def-hoja c3)
  (def-hoja c4)

  (def-hoja m1)
  (def-hoja m2)
  (def-hoja m3)
  (def-hoja m4)

  (def-hoja d1)
  (def-hoja d2)
  (def-hoja d3)
  (def-hoja d4))

(defparameter libro-matcom (libro :hojas (list
                                          ;; c1
                                          (hoja-c1 "11")
                                          (hoja-c1 "12")
                                          (hoja-c1 "13")
                                          (hoja-c1 "21")
                                          (hoja-c1 "22")
                                          ;; c2
                                          (hoja-c2 "11")
                                          (hoja-c2 "12")
                                          (hoja-c2 "13")
                                          ;; c3
                                          (hoja-c3 "11")
                                          (hoja-c3 "12")
                                          ;; c4
                                          (hoja-c4 "11")
                                          (hoja-c4 "12"))))

(defclass output-org-mode () ())
(defparameter org (make-instance 'output-org-mode))

(defun write-org-table-row (&key cols row-contents stream)
     "Writes an empty row with the given number of cols,
or if row-contents is a list writes a row with those elements.
There is not error check, so make sure that the length of the list matches the desired number of columns.

If row-contents is non-nil, then the value of cols is ignored."
   (if row-contents ;; let's add those contents
       (progn
         ;; first, the initial |
         (format stream "| ")
         (format stream "~{ ~a |~}"
                 row-contents))
       ;; else  
       (progn
         (loop for i from 1 to cols
               do (format stream "| "))
         ;; the final |
         (format stream "|"))))

(defun draw-empty-table (&key rows row-names cols (cell-size 1) stream add-line-after-row)
  "Draws an empty table with the give number of rows and cols"
  (setf rows (if row-names (length row-names)
                 rows))
  (loop for i from 1 to rows
          do ;; let's add as many rows as cell-size
             (loop for j from 1 to cell-size
                   doing
                      ;; if it is the first row of the cell add the row-name
                      ;; if exists
                      (if (= j 1)
                          (progn
                            ;; the initial |
                            (format stream "| ")
                            (loop for k from 1 to cols
                                  doing
                                     (if (and (= k 1)
                                              row-names)
                                         (format stream "~a |" (nth (1- i) row-names))
                                         ;; else
                                         (format stream " | "))
                                  finally (format stream "~%")))
                          ;; else
                          (progn
                            (write-org-table-row :cols cols :stream stream)
                            (format stream "~%"))))
             ;; the newline after the row
             (when add-line-after-row
    	(format stream "|-~%"))))

(defun draw-table-with-content (&key (cell-size 1) stream add-line-after-row row-contents)
  "Draws a table with the given content.
The content includes the name of the row.
Where cell-size is more than one, add content in the first row of the cell."
  ;; let's iterate over the contents
  (loop for row in row-contents
          do ;; let's add as many rows as cell-size
             (loop for j from 1 to cell-size
                   doing
                      ;; if it is the first row of the cell add the content
                      (if (= j 1)
                          (format stream "| ~{~a |~} ~%"
                                  row)
                          ;; else
                          ;; just draw a |
                          (progn
                            (format stream "|~%"))))
             ;; the newline after the row (cell)
             (when add-line-after-row
       	(format stream "|-~%"))))

(defun make-org-table (&key
                         cols
                         col-names
                         (cell-size 1)
                         (rows 3)
                         row-names
                         row-contents 
                         stream
                         add-line-after-row
                         add-initial-line
                         )
  "Writes the given table in org-mode.
If table-contents is non nil, it should be a list of lists, where each sublist represents a row.
The elements of table-contents should not include the columns-name row.
There is not error-checking, so make sure all the elements in the list have the same length."
  (let* ((actual-column-names (if row-contents (first row-contents)
                                  ;; else
                                  (if col-names
                                      (progn
                                        (setf cols (length col-names))
                                        col-names)
                                      ;; else
                                         (if cols
                                             (make-list cols :initial-element " ")
                                             (error "In make-org-table either cols or col-names is required")))))
         (cols (or cols
                   (length col-names)))
         (row-names (if row-names
                        (progn
                          (setf rows (length row-names))
                          row-names)
                        ;; else
                        (if rows
                            (make-list rows :initial-element " ")
                            (error "In make-org-table either rows or row-names is required"))))
         (rows (or rows
                   (length row-names))))
    ;; let's add a header line
    (if add-initial-line
        (format stream "|-~%"))
    ;; the following code writes the heading row
    ;; if there are col-names AND there are not row-contents
    (if (and
         col-names
         (not row-contents))
        (progn
          (loop for col-name in actual-column-names
                doing (format stream "| ~a " col-name))
          ;; now the final |
          (format stream "|~%")
          ;; let's add a header separator
          (format stream "|-~%")))

    ;; now, let's choose acoording to table-contents
    (if row-contents
        (draw-table-with-content :cell-size cell-size
                                 :stream stream
                                 :add-line-after-row add-line-after-row
                                 :row-contents row-contents)
        ;; else
        (draw-empty-table :rows rows
                          :row-names row-names
                          :cols cols
                          :cell-size cell-size
                          :stream stream
                          :add-line-after-row add-line-after-row))))

(defmethod generate-code ((node clase-tabla)
       		   (lang output-org-mode)
       		   (stream t))
  (make-org-table :stream stream
       	   :cols (columnas node)
       	   :rows (filas node)
       	   :add-line-after-row t
                  :add-initial-line t
                  :row-contents (contenido-de-la-tabla node)
       	   :row-names (nombres-filas node)
       	   :col-names (nombres-columnas node)))

(defmethod generate-code ((node clase-hoja)
      		    (lang output-org-mode)
      		    (stream t))
  (format stream "* ~a~2%" (grupo node))

  (format stream "Horario: ~%")

  (generate-code (horario node) lang stream)

  (format stream "~2%")

  (format stream "Asignaturas: ~%")

  (generate-code (asignaturas node) lang stream))

(defmethod generate-code ((node clase-libro)
      		    (lang output-org-mode)
      		    (stream t))
  (loop for h in (hojas node)
        doing 
           (generate-code h lang stream)
           (format stream "~2%")))
