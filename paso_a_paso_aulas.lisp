;; paso_a_paso_aulas.lisp
;; Este archivo inicializa el modelo y genera las llamadas integrándose a hoja_con_formulas.py

(load "codigo-tesis.lisp")

;; 1. Usar las macros de codigo-tesis.lisp para definir la nueva clase hoja-aulas
(defclass* hoja-aulas ()
  (lunes
   martes
   miercoles
   jueves
   viernes))

;; 2. Definir una nueva clase de dispatch para esta generación con dicionario Python
(defclass output-python-config () ())
(defparameter python-config (make-instance 'output-python-config))

;; Un helper para formatear listas Lisp como arrays de python
(defun format-python-list (lst stream)
  (if (null lst)
      (format stream "None")
      (progn
        (format stream "[")
        (loop for x in lst
              for i from 0
              do
              (if (> i 0) (format stream ", "))
              (if (listp x)
                  (format-python-list x stream)
                  (cond ((numberp x) (format stream "~a" x))
                        ((equal x "") (format stream "\"\""))
                        (t (format stream "\"~a\"" x)))))
        (format stream "]"))))

;; 3. generate-code para las tablas
;; Mapea todos los datos para la llave "data" del diccionario
(defmethod generate-code ((node clase-tabla)
                          (lang output-python-config)
                          (stream t))
  ;; Si NO es Lunes, necesitamos añadir un margen y la fila de nombres_columnas
  (if (not (string-equal (id node) "Lunes"))
      (progn
        (format stream "data_aulas.append([])~%")
        (format stream "data_aulas.append(")
        (format-python-list (nombres-columnas node) stream)
        (format stream ")~%")))
        
  ;; Siempre añadimos todas las filas de la tabla iterativamente
  (loop for r-idx from 0 to (1- (filas node))
        do
        (format stream "data_aulas.append(")
        (let ((row (list (nth r-idx (nombres-filas node)))))
          (if (contenido-de-la-tabla node)
              (setf row (append row (nth r-idx (contenido-de-la-tabla node))))
              (setf row (append row (make-list (- (columnas node) 1) :initial-element ""))))
          (format-python-list row stream))
        (format stream ")~%")))

;; 4. generate-code para la hoja-aulas 
;; Escribe la importación del script del usuario, agrupa las tablas y lo llama
(defmethod generate-code ((node clase-hoja-aulas)
                          (lang output-python-config)
                          (stream t))
  (format stream "import sys~%")
  (format stream "from hoja_con_formulas import generar_excel_personalizado~%~%")
  
  (format stream "data_aulas = []~%")
  (format stream "headers_aulas = ")
  (format-python-list (nombres-columnas (lunes node)) stream)
  (format stream "~%~%")
  
  (generate-code (lunes node) lang stream)
  (generate-code (martes node) lang stream)
  (generate-code (miercoles node) lang stream)
  (generate-code (jueves node) lang stream)
  (generate-code (viernes node) lang stream)
  
  (format stream "~%config_excel = {~%")
  (format stream "    'sheets': [~%")
  (format stream "        {~%")
  (format stream "            'title': 'Aulas',~%")
  (format stream "            'headers': headers_aulas,~%")
  (format stream "            'data': data_aulas,~%")
  (format stream "            'column_widths': {i: 12 for i in range(1, 12)},~%")
  (format stream "            'header_style': {'bold': True, 'align': 'center', 'bg_color': 'F09E9E'}~%")
  (format stream "        }~%")
  (format stream "    ]~%")
  (format stream "}~%~%")
  (format stream "generar_excel_personalizado(config_excel, 'Aulas_Testing.xlsx')~%"))


;; ===============================================
;; Setup Local para Ejecutar el test
;; ===============================================

(defun crear-tabla-dia (dia id-tabla contenido)
  (tabla :id id-tabla
         :filas 6
         :columnas 11
         :nombres-columnas (list dia "Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5" "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab")
         :nombres-filas '("1ro" "2do" "3ro" "4to" "5to" "6to")
         :contenido-de-la-tabla contenido))

(defparameter tbl-lunes (crear-tabla-dia "Lunes" "Lunes" 
                       '(("" "" "" "" "" "C111,C112" "" "" "" "C113")
                         ("" "" "" "" "C112" "C113" "" "" "" "")
                         ("" "" "" "" "" "C111,C113" "" "C112" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-martes (crear-tabla-dia "Martes" "Martes" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-miercoles (crear-tabla-dia "Miércoles" "Miércoles" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-jueves (crear-tabla-dia "Jueves" "Jueves" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter tbl-viernes (crear-tabla-dia "Viernes" "Viernes" 
                       '(("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" "")
                         ("" "" "" "" "" "" "" "" "" ""))))

(defparameter mi-hoja-aulas (hoja-aulas :lunes tbl-lunes
                                        :martes tbl-martes
                                        :miercoles tbl-miercoles
                                        :jueves tbl-jueves
                                        :viernes tbl-viernes))

(with-open-file (f "ejecutar_hoja_con_formulas.py" :direction :output :if-exists :supersede)
  (generate-code mi-hoja-aulas python-config f))

(format t "¡Fichero Python 'ejecutar_hoja_con_formulas.py' generado con éxito!~%")
