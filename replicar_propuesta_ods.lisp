;; replicar_propuesta_ods.lisp
;; Reconstruye el libro del ODS usando las macros/clases de codigo-tesis.lisp
;; y genera un script Python compatible con hoja_con_formulas.py.

(load "codigo-tesis.lisp")
(load "variables_horario.lisp")

(defclass* hoja-aulas ()
  (lunes
   martes
   miercoles
   jueves
   viernes
   grupos))

(defclass* output-python-config ()
  (group-title-prefix
   group-header-prefix-cells
   group-header-gap-cells
   group-header-suffix-cells
   group-row-prefix-empty-cells
   group-row-suffix-prefix-cells
   group-resaltar-label
   group-aulas-lateral-values
   group-column-width
   group-table-ranges
   group-table-block-ranges
   group-merge-ranges
   group-table-borders
   group-default-border-color
   group-default-border-style
   group-extra-range-styles
   aulas-title
   aulas-column-width
   aulas-range-styles
   aulas-header-style
   aulas-table-ranges
  aulas-table-borders
   aulas-border-color
   aulas-border-style
   aulas-header-row-prefix-cells
   aulas-column-index-labels
   aulas-row-prefix-empty-cells
   python-module
   python-function
   output-excel-file
   horario-cantidad-turnos
   horario-alto-celda
   horario-ancho-celda
  horario-first-column-label
  horario-id-prefix
  asignaturas-id-prefix
   horario-border-color
   horario-border-style
   horario-range-styles
   asignaturas-alto-celda
   asignaturas-ancho-celda
   asignaturas-headers
   aulas-column-names
   aulas-row-names))

(defun write-python-value (value stream)
  (cond ((null value)
         (format stream "None"))
        ((stringp value)
         (format stream "~s" value))
        ((numberp value)
         (format stream "~a" value))
        ((listp value)
         (format stream "[")
         (loop for x in value
               for i from 0
               do (when (> i 0)
                    (format stream ", "))
                  (write-python-value x stream))
         (format stream "]"))
        (t
         (format stream "~s" value))))

(defun string-replace-all (string old new)
  "Reemplaza todas las ocurrencias de OLD con NEW en STRING."
  (let ((result string)
        (old-len (length old)))
    (loop for pos = (search old result)
          while pos
          do (setf result (concatenate 'string
                                       (subseq result 0 pos)
                                       new
                                       (subseq result (+ pos old-len)))))
    result))

(defun substitute-ranges (formula asig-range aulas-range)
  "Reemplaza {asignaturas_abrev_range} y {aulas_range} en la fórmula."
  (let ((result formula))
    (setf result (string-replace-all result "{asignaturas_abrev_range}" 
                                     (format nil "$~a" asig-range)))
    (setf result (string-replace-all result "{aulas_range}" 
                                     (format nil "$~a" aulas-range)))
    result))

(defun substitute-asig-refs (formula faltan-col frec-col asignadas-col)
  "Reemplaza {faltan_celda}, {frec_celda}, {asignadas_celda} con referencias dinámicas."
  (let ((result formula))
    (setf result (string-replace-all result "{faltan_celda}" 
                                     (format nil "~a{fila}" faltan-col)))
    (setf result (string-replace-all result "{frec_celda}" 
                                     (format nil "~a{fila}" frec-col)))
    (setf result (string-replace-all result "{asignadas_celda}" 
                                     (format nil "~a{fila}" asignadas-col)))
    result))

(defun generar-reglas-formato-condicional (stream horario-range asig-abrev-range aulas-range
                                            asig-col asig-start-row asig-end-row
                                            faltan-col frec-col asignadas-col
                                            horario-row-step)
  "Genera las reglas de formato condicional en formato Python.
   Reemplaza los placeholders dinámicos en las fórmulas."
  (format stream "    'conditional_format_rules': [~%")
  
  ;; Reglas para el horario (asignaturas y aulas)
  (loop for regla in *cfg-formato-condicional-reglas*
        for i from 0
        do (let* ((tipo (getf regla :tipo))
                   (formula-template (getf regla :formula))
                   (color-var (getf regla :color-var))
                   (color (symbol-value color-var))
                   (aplicar-a (getf regla :aplicar-a))
                   (row-offset (cond ((string= tipo "filas_pares") 0)
                                     ((string= tipo "filas_impares") 1)
                                     (t nil)))
                   ;; Reemplazar rangos dinámicos en la fórmula
                   (formula (substitute-ranges formula-template 
                                               asig-abrev-range 
                                               aulas-range)))
              (when (> i 0) (format stream ",~%"))
              (format stream "        {~%")
              (format stream "            'tipo': '~a',~%" tipo)
              (format stream "            'rango': '~a',~%" horario-range)
              (format stream "            'formula': '~a',~%" formula)
              (format stream "            'color': '~a'" color)
              (when (member tipo '("filas_pares" "filas_impares" "pares_con_siguiente") :test #'string=)
                (format stream ",~%            'row_step': ~a" horario-row-step))
              (when row-offset
                (format stream ",~%            'row_start_offset': ~a" row-offset))
              (when (string= tipo "pares_con_siguiente")
                (format stream ",~%            'next_offset': ~a"
                        (if (> horario-row-step 1) 1 0)))
              (when aplicar-a
                (format stream ",~%            'aplicar_a': '~a'" aplicar-a))
              (format stream "~%        }")))
  
  ;; Reglas para la tabla de asignaturas
  (let ((asig-range (format nil "~a~a:~a~a" asig-col asig-start-row asig-col asig-end-row)))
    (loop for regla in *cfg-formato-condicional-asignaturas*
          do (let* ((tipo (getf regla :tipo))
                    (formula-template (getf regla :formula))
                    (color-var (getf regla :color-var))
                    (color (symbol-value color-var))
                    ;; Reemplazar referencias a columnas dinámicas
                    (formula (substitute-asig-refs formula-template 
                                                   faltan-col frec-col asignadas-col)))
               (format stream ",~%        {~%")
               (format stream "            'tipo': '~a',~%" tipo)
               (format stream "            'rango': '~a',~%" asig-range)
               (format stream "            'formula': '~a',~%" formula)
               (format stream "            'color': '~a'~%" color)
               (format stream "        }"))))
  
  (format stream "~%    ]"))

(defun make-blank-row (size)
  (make-list size :initial-element ""))

(defun insert-at-index (items index new-item)
  "Inserta NEW-ITEM en ITEMS antes de INDEX."
  (let ((safe-index (max 0 (min index (length items)))))
    (append (subseq items 0 safe-index)
            (list new-item)
            (subseq items safe-index))))

(defun move-last-item-to-index (items index)
  "Mueve el último elemento de ITEMS antes de INDEX."
  (let* ((len (length items))
         (safe-index (max 0 (min index (max 0 (1- len))))))
    (if (<= len 1)
        items
        (let ((last-item (nth (1- len) items))
              (head (subseq items 0 (1- len))))
          (append (subseq head 0 safe-index)
                  (list last-item)
                  (subseq head safe-index))))))

(defparameter *headers-horario* '("Lunes" "Martes" "Miércoles" "Jueves " "Viernes"))
(defparameter *rows-horario*
  '("Turno 1" ""
    "Turno 2" ""
    "Turno 3" ""
    "Turno 4" ""
    "Turno 5" ""
    "Turno 6" ""))
(defparameter *headers-asignaturas* '("Abrev" "Asignaturas" "Frec" "Faltan" "Asignadas"))
(defparameter *aulas-lateral* '("1" "2" "3" "4" "5" "6" "7" "8" "9" "Lab" "" "" "" ""))

;; Colores configurables (RGB en hex, sin #)
(defparameter *color-borde-tablas* "B3B3B3")   ;; gris tenue un poco mas oscuro
(defparameter *color-verde-asignaturas* "A9D18E")
(defparameter *color-rojo-aulas* "E6B8AF")
(defparameter *color-verde-horario-encabezado* "C6EFCE")
(defparameter *color-rojo-turnos* "F4CCCC")
(defparameter *color-borde-horario-aula* "4F81BD")

;; Colores para formato condicional
(defparameter *color-cf-asignatura-invalida* "F4A460")  ;; salmón
(defparameter *color-cf-aula-invalida* "FFD700")        ;; amarillo
(defparameter *color-cf-asignatura-sin-aula* "FF0000") ;; rojo
(defparameter *color-cf-asignatura-completa* "00FF00") ;; verde
(defparameter *color-cf-asignatura-excedida* "FF6B6B") ;; rojo claro
(defparameter *color-cf-asignatura-parcial* "FFA500")  ;; naranja

;; Plantillas de reglas de formato condicional
;; Cada regla tiene: tipo, formula (con placeholders), color
;; Placeholders: {celda}, {celda_siguiente}, {fila}, {columna}
;; Los rangos dinámicos se insertan en el momento de generar

(defparameter *cfg-formato-condicional-reglas*
  '(;; Formato 1: Asignaturas no válidas en filas pares del horario
    (:id "asignatura-invalida"
     :tipo "filas_pares"
     :formula "AND({celda}<>\"\", COUNTIF({asignaturas_abrev_range},{celda})=0)"
     :color-var *color-cf-asignatura-invalida*
     :descripcion "Asignatura no está en la lista de abreviaturas válidas")
    
    ;; Formato 2: Aulas no válidas en filas impares del horario
    (:id "aula-invalida"
     :tipo "filas_impares"
     :formula "AND({celda}<>\"\", COUNTIF({aulas_range},{celda})=0)"
     :color-var *color-cf-aula-invalida*
     :descripcion "Aula no está en la lista de aulas válidas")
    
    ;; Formato 3: Asignatura sin aula asignada
    (:id "asignatura-sin-aula"
     :tipo "pares_con_siguiente"
     :formula "AND({celda}<>\"\", {celda_siguiente}=\"\")"
     :color-var *color-cf-asignatura-sin-aula*
     :aplicar-a "siguiente"
     :descripcion "Asignatura tiene contenido pero aula está vacía")))

;; Reglas para la tabla de asignaturas (columna J según valores de L, K, M)
(defparameter *cfg-formato-condicional-asignaturas*
  '(;; Verde: Asignatura completa (Faltan = 0 y Asignadas > 0)
    (:id "asignatura-completa"
     :tipo "rango"
     :formula "AND({celda}<>\"\", {asignadas_celda}>0, {faltan_celda}=0)"
     :color-var *color-cf-asignatura-completa*
     :descripcion "Asignatura completamente asignada")
    
    ;; Rojo: Asignatura excedida (Faltan < 0)
    (:id "asignatura-excedida"
     :tipo "rango"
     :formula "AND({celda}<>\"\", {asignadas_celda}>0, {faltan_celda}<0)"
     :color-var *color-cf-asignatura-excedida*
     :descripcion "Asignatura con más horas de las necesarias")
    
    ;; Naranja: Asignatura parcial (0 < Faltan < Frec)
    (:id "asignatura-parcial"
     :tipo "rango"
     :formula "AND({celda}<>\"\", {asignadas_celda}>0, {faltan_celda}>0, {faltan_celda}<{frec_celda})"
     :color-var *color-cf-asignatura-parcial*
     :descripcion "Asignatura parcialmente asignada")))

;; =============================================================================
;; CONFIGURACIÓN DE FÓRMULAS DE CELDAS
;; =============================================================================
;; Fórmulas configurables que se insertan en celdas específicas.
;; Cada fórmula tiene:
;;   :id - Identificador único
;;   :celda - Posición relativa donde insertar (ej: "N{asig-end+1}" o "I{asig-end+2}")
;;   :formula - Plantilla de fórmula Excel con placeholders
;;   :etiqueta - Celda donde poner la etiqueta (opcional)
;;   :etiqueta-texto - Texto de la etiqueta (opcional)
;;
;; Placeholders disponibles:
;;   {asig-start} - Fila inicio asignaturas (ej: 4)
;;   {asig-end} - Fila fin asignaturas
;;   {aulas-start} - Fila inicio aulas (ej: 4)
;;   {aulas-end} - Fila fin aulas
;;   {horario-start} - Fila inicio horario (ej: 4)
;;   {horario-end} - Fila fin horario

(defparameter *cfg-formulas-hoja*
  '(;; Total de asignaturas (cuenta celdas no vacías en columna Abrev)
    (:id "total-asignaturas"
     :celda-fila "{asig-end+1}"
     :celda-col "M"
     :formula "=COUNTA(I{asig-start}:I{asig-end})"
     :etiqueta-col "L"
     :etiqueta-texto "Total:")
    
    ;; Suma de frecuencias
    (:id "suma-frecuencias"
     :celda-fila "{asig-end+2}"
     :celda-col "M"
     :formula "=SUM(K{asig-start}:K{asig-end})"
     :etiqueta-col "L"
     :etiqueta-texto "Σ Frec:")
    
    ;; Total de aulas únicas
    (:id "total-aulas"
     :celda-fila "{aulas-end+1}"
     :celda-col "O"
     :formula "=COUNTA(O{aulas-start}:O{aulas-end})"
     :etiqueta-col "N"
     :etiqueta-texto "Total:")
    
     ;; Turnos asignados (cuenta celdas no vacías en horario)
    (:id "turnos-ocupados"
     :celda-fila "{horario-end+1}"
     :celda-col "G" 
     :formula "=COUNTA(C{horario-start}:G{horario-end})/{horario-row-step}"
     :etiqueta-col "F"
     :etiqueta-texto "Ocupados:")))

(defun substituir-placeholders-formula (template asig-start asig-end aulas-start aulas-end horario-start horario-end horario-row-step)
  "Reemplaza los placeholders en una plantilla de fórmula."
  (let ((result template))
    (setf result (string-replace-all result "{asig-start}" (write-to-string asig-start)))
    (setf result (string-replace-all result "{asig-end}" (write-to-string asig-end)))
    (setf result (string-replace-all result "{aulas-start}" (write-to-string aulas-start)))
    (setf result (string-replace-all result "{aulas-end}" (write-to-string aulas-end)))
    (setf result (string-replace-all result "{horario-start}" (write-to-string horario-start)))
    (setf result (string-replace-all result "{horario-end}" (write-to-string horario-end)))
    (setf result (string-replace-all result "{horario-row-step}" (write-to-string (max 1 horario-row-step))))
    result))

(defun calcular-fila-formula (fila-spec asig-start asig-end aulas-start aulas-end horario-start horario-end)
  "Calcula la fila numérica desde una especificación como '{asig-end+1}'."
  (cond
    ;; Número directo
    ((numberp fila-spec) fila-spec)
    ;; String con placeholder
    ((stringp fila-spec)
     (let ((spec fila-spec))
       ;; Parsear formatos como "{asig-end+1}" o "{aulas-end}"
       (cond
         ((search "{asig-end+" spec)
          (let* ((offset-str (subseq spec (+ (search "{asig-end+" spec) 10)))
                 (offset (parse-integer (string-trim "}" offset-str))))
            (+ asig-end offset)))
         ((search "{asig-end}" spec) asig-end)
         ((search "{aulas-end+" spec)
          (let* ((offset-str (subseq spec (+ (search "{aulas-end+" spec) 11)))
                 (offset (parse-integer (string-trim "}" offset-str))))
            (+ aulas-end offset)))
         ((search "{aulas-end}" spec) aulas-end)
         ((search "{horario-end+" spec)
          (let* ((offset-str (subseq spec (+ (search "{horario-end+" spec) 13)))
                 (offset (parse-integer (string-trim "}" offset-str))))
            (+ horario-end offset)))
         ((search "{horario-end}" spec) horario-end)
         (t (parse-integer spec)))))
    (t 1)))

(defun generar-formulas-hoja (stream asig-start asig-end aulas-start aulas-end horario-start horario-end horario-row-step)
  "Genera el array de fórmulas para una hoja."
  (format stream "    'formulas': [~%")
  (loop for formula-cfg in *cfg-formulas-hoja*
        for i from 0
        do (let* ((fila-spec (getf formula-cfg :celda-fila))
                  (col (getf formula-cfg :celda-col))
                  (formula-template (getf formula-cfg :formula))
                  (etiqueta-col (getf formula-cfg :etiqueta-col))
                  (etiqueta-texto (getf formula-cfg :etiqueta-texto))
                  (fila (calcular-fila-formula fila-spec asig-start asig-end 
                                               aulas-start aulas-end 
                                               horario-start horario-end))
                  (formula (substituir-placeholders-formula formula-template 
                                                           asig-start asig-end 
                                                           aulas-start aulas-end 
                                                           horario-start horario-end
                                                           horario-row-step))
                  (col-num (1+ (- (char-code (char col 0)) (char-code #\A)))))
             ;; Etiqueta
             (when etiqueta-texto
               (let ((etiqueta-col-num (1+ (- (char-code (char etiqueta-col 0)) (char-code #\A)))))
                 (when (> i 0) (format stream ",~%"))
                 (format stream "        {'row': ~a, 'col': ~a, 'value': ~s}" 
                         fila etiqueta-col-num etiqueta-texto)
                 (format stream ",~%")))
             ;; Fórmula
             (unless etiqueta-texto
               (when (> i 0) (format stream ",~%")))
             (format stream "        {'row': ~a, 'col': ~a, 'value': ~s}" 
                     fila col-num formula)))
   (format stream "~%    ],~%"))

(defun generar-formulas-dinamicas-hoja (stream asig-start asig-end frec-col faltan-col asignadas-col abrev-col)
  "Genera el bloque Python para fórmulas dinámicas de Faltan y Asignadas por fila de asignatura."
  (let ((frec-col-num (1+ (- (char-code (char frec-col 0)) (char-code #\A))))
        (faltan-col-num (1+ (- (char-code (char faltan-col 0)) (char-code #\A))))
        (asignadas-col-num (1+ (- (char-code (char asignadas-col 0)) (char-code #\A))))
        (abrev-col-num (1+ (- (char-code (char abrev-col 0)) (char-code #\A)))))
    (format stream "~%# Generar fórmulas dinámicas de Faltan y Asignadas para cada hoja de grupo~%")
    (format stream "for sheet in sheets_cfg:~%")
    (format stream "    if sheet.get('title') == 'Aulas':~%")
    (format stream "        continue~%")
    (format stream "    data = sheet.get('data', [])~%")
    (format stream "    horario_range = sheet.get('horario_data_range', '$C$4:$G$15')~%")
    (format stream "    formulas = sheet.get('formulas', [])~%")
    (format stream "    for row_idx, row_data in enumerate(data):~%")
    (format stream "        abrev = row_data[~a] if len(row_data) > ~a else None~%" 
            (1- abrev-col-num) (1- abrev-col-num))
    (format stream "        if abrev and isinstance(abrev, str) and abrev.strip() and abrev.strip() != 'Abrev':~%")
    (format stream "            excel_row = row_idx + 1~%")
    (format stream "            formulas.append({'row': excel_row, 'col': ~a, 'value': f'=COUNTIF({horario_range},I{excel_row})'})~%"
            asignadas-col-num)
    (format stream "            formulas.append({'row': excel_row, 'col': ~a, 'value': f'=~a{excel_row}-~a{excel_row}'})~%"
            faltan-col-num frec-col asignadas-col)
    (format stream "    sheet['formulas'] = formulas~%~%")))

(defparameter *rangos-merge-columna-turnos*
  '("B4:B5"
    "B6:B7"
    "B8:B9"
    "B10:B11"
    "B12:B13"
    "B14:B15"))

(defparameter *cfg-group-title-prefix* "Grupo ")
(defparameter *cfg-group-header-prefix-cells* '("" ""))
(defparameter *cfg-group-header-gap-cells* '(""))
(defparameter *cfg-group-header-suffix-cells* '("" "Aulas"))
(defparameter *cfg-group-row-prefix-empty-cells* '(""))
(defparameter *cfg-group-row-suffix-prefix-cells* '(""))
(defparameter *cfg-group-resaltar-label* "Resaltar")
(defparameter *cfg-group-aulas-lateral* *aulas-lateral*)
(defparameter *cfg-group-column-width* 14)
(defparameter *cfg-group-table-ranges*
  '("B3:G3" "B4:B15" "C4:G15" "I3:M3" "I4:M15" "N3:N15"))
(defparameter *cfg-group-table-block-ranges*
  '(:horario-header "B3:G3"
    :turnos "B4:B15"
    :horario "C4:G15"
    :asig-header "I3:L3"
    :asig "I4:L15"
    :aulas "N3:N15"))
(defparameter *cfg-group-merge-ranges* *rangos-merge-columna-turnos*)
(defparameter *cfg-group-table-borders* t)
(defparameter *cfg-group-default-border-color* *color-borde-tablas*)
(defparameter *cfg-group-default-border-style* "medium")
(defparameter *cfg-group-extra-range-styles*
  (list (list "I3:I15" *color-verde-asignaturas*)))

(defparameter *cfg-aulas-title* "Aulas")
(defparameter *cfg-aulas-column-width* 12)
(defparameter *cfg-aulas-range-styles*
  (list (list "B3:L3" *color-rojo-aulas*)
        (list "B11:L11" *color-rojo-aulas*)
        (list "B19:L19" *color-rojo-aulas*)
        (list "B27:L27" *color-rojo-aulas*)
        (list "B35:L35" *color-rojo-aulas*)))
(defparameter *cfg-aulas-header-style*
  (list :bold t :align "center" :bg-color *color-rojo-aulas*))
(defparameter *cfg-aulas-table-ranges*
  '("B3:L9" "B11:L17" "B19:L25" "B27:L33" "B35:L41"))
(defparameter *cfg-aulas-table-borders* t)
(defparameter *cfg-aulas-border-color* *color-borde-tablas*)
(defparameter *cfg-aulas-border-style* "thick")
(defparameter *cfg-aulas-header-row-prefix-cells* '("" ""))
(defparameter *cfg-aulas-column-index-labels*
  '("1" "2" "3" "4" "5" "6" "7" "8" "9" "Lab"))
(defparameter *cfg-aulas-row-prefix-empty-cells* '(""))

(defparameter *cfg-python-module* "hoja_con_formulas")
(defparameter *cfg-python-function* "generar_excel_personalizado")
(defparameter *cfg-output-excel-file* "propuesta_horarios_desde_lisp.xlsx")

(defparameter *cfg-horario-cantidad-turnos* 6)
(defparameter *cfg-horario-alto-celda* 3)
(defparameter *cfg-horario-ancho-celda* 1)
(defparameter *cfg-horario-first-column-label* "Turnos")
(defparameter *cfg-horario-id-prefix* "horario_")
(defparameter *cfg-asignaturas-id-prefix* "asignaturas_")
(defparameter *cfg-horario-border-color* *color-borde-horario-aula*)
(defparameter *cfg-horario-border-style* "medium")
(defparameter *cfg-horario-range-styles*
  (list (list "B3:G3" *color-verde-horario-encabezado*)
        (list "B4:B15" *color-rojo-turnos*)))

(defparameter *cfg-asignaturas-alto-celda* 1)
(defparameter *cfg-asignaturas-ancho-celda* 1)
(defparameter *cfg-asignaturas-headers* *headers-asignaturas*)

(defparameter *cfg-aulas-column-names*
  '("Aula 1" "Aula 2" "Aula 3" "Aula 4" "Aula 5"
    "Aula 6" "Aula 7" "Aula 8" "Aula 9" "Lab"))
(defparameter *cfg-aulas-row-names*
  '("1ro" "2do" "3ro" "4to" "5to" "6to"))

(defparameter python-config
  (output-python-config
    :group-title-prefix *cfg-group-title-prefix*
    :group-header-prefix-cells *cfg-group-header-prefix-cells*
    :group-header-gap-cells *cfg-group-header-gap-cells*
    :group-header-suffix-cells *cfg-group-header-suffix-cells*
    :group-row-prefix-empty-cells *cfg-group-row-prefix-empty-cells*
    :group-row-suffix-prefix-cells *cfg-group-row-suffix-prefix-cells*
    :group-resaltar-label *cfg-group-resaltar-label*
    :group-aulas-lateral-values *cfg-group-aulas-lateral*
    :group-column-width *cfg-group-column-width*
    :group-table-ranges *cfg-group-table-ranges*
    :group-table-block-ranges *cfg-group-table-block-ranges*
    :group-merge-ranges *cfg-group-merge-ranges*
    :group-table-borders *cfg-group-table-borders*
    :group-default-border-color *cfg-group-default-border-color*
    :group-default-border-style *cfg-group-default-border-style*
    :group-extra-range-styles *cfg-group-extra-range-styles*
    :aulas-title *cfg-aulas-title*
    :aulas-column-width *cfg-aulas-column-width*
    :aulas-range-styles *cfg-aulas-range-styles*
    :aulas-header-style *cfg-aulas-header-style*
    :aulas-table-ranges *cfg-aulas-table-ranges*
    :aulas-table-borders *cfg-aulas-table-borders*
    :aulas-border-color *cfg-aulas-border-color*
    :aulas-border-style *cfg-aulas-border-style*
    :aulas-header-row-prefix-cells *cfg-aulas-header-row-prefix-cells*
    :aulas-column-index-labels *cfg-aulas-column-index-labels*
    :aulas-row-prefix-empty-cells *cfg-aulas-row-prefix-empty-cells*
    :python-module *cfg-python-module*
    :python-function *cfg-python-function*
    :output-excel-file *cfg-output-excel-file*
    :horario-cantidad-turnos *cfg-horario-cantidad-turnos*
    :horario-alto-celda *cfg-horario-alto-celda*
    :horario-ancho-celda *cfg-horario-ancho-celda*
    :horario-first-column-label *cfg-horario-first-column-label*
    :horario-id-prefix *cfg-horario-id-prefix*
    :asignaturas-id-prefix *cfg-asignaturas-id-prefix*
    :horario-border-color *cfg-horario-border-color*
    :horario-border-style *cfg-horario-border-style*
    :horario-range-styles *cfg-horario-range-styles*
    :asignaturas-alto-celda *cfg-asignaturas-alto-celda*
    :asignaturas-ancho-celda *cfg-asignaturas-ancho-celda*
    :asignaturas-headers *cfg-asignaturas-headers*
    :aulas-column-names *cfg-aulas-column-names*
    :aulas-row-names *cfg-aulas-row-names*))

(defun construir-filas-hoja-grupo (node &optional (config python-config))
  (let* ((horario-tabla (horario node))
         (asig-tabla (asignaturas node))
         (horario-columnas (nombres-columnas horario-tabla))
         (horario-columnas-visibles
           (if (and horario-columnas
                    (string= (first horario-columnas)
                             (horario-first-column-label config)))
               (rest horario-columnas)
               horario-columnas))
         (horario-rows (contenido-de-la-tabla horario-tabla))
         (asig-rows (if asig-tabla (contenido-de-la-tabla asig-tabla) '()))
         (header-prefix (group-header-prefix-cells config))
         (header-gap (group-header-gap-cells config))
         (header-suffix (group-header-suffix-cells config))
         (asig-columnas (if asig-tabla (nombres-columnas asig-tabla) '()))
         (header-row (append header-prefix
                             horario-columnas-visibles
                             header-gap
                             asig-columnas
                             header-suffix))
         (row-length (length header-row))
         (row-prefix-empty (group-row-prefix-empty-cells config))
         (row-suffix-prefix (group-row-suffix-prefix-cells config))
         (resaltar-label (group-resaltar-label config))
         ;; Extraer aulas únicas del horario y crear lista dinámica
         (aulas-unicas (extraer-aulas-unicas horario-rows))
         (total-rows (length (nombres-filas horario-tabla)))
         (aulas-lateral-dinamica 
           (append aulas-unicas 
                   (make-list (max 0 (- total-rows (length aulas-unicas))) 
                              :initial-element "")))
         (resaltar-idx (max 0 (1- total-rows)))
         (title-prefix (list (group-title-prefix config) (grupo node)))
         (result (list
            (append title-prefix
              (make-blank-row (- row-length (length title-prefix))))
            (make-blank-row row-length)
            header-row)))
    (loop for i from 0 below total-rows
          do
             (let* ((fila-horario (if (< i (length horario-rows))
                                      (nth i horario-rows)
                                      (make-blank-row (length horario-columnas-visibles))))
                    (fila-asig (if asig-tabla
                                   (cond ((< i (length asig-rows)) (nth i asig-rows))
                                         (t (make-blank-row (length (nombres-columnas asig-tabla)))))
                                   '()))
                    (aula-valor (if (< i (length aulas-lateral-dinamica))
                                    (nth i aulas-lateral-dinamica)
                                    ""))
                    (fila-completa (append
                                    (append row-prefix-empty
                                            (list (nth i (nombres-filas horario-tabla))))
                                    fila-horario
                                    header-gap
                                    fila-asig
                                    (append row-suffix-prefix (list aula-valor)))))
               (setf result (append result (list fila-completa)))))
    result))

(defun construir-filas-aulas (node &optional (config python-config))
  (let ((dias (list (lunes node)
                    (martes node)
                    (miercoles node)
                    (jueves node)
                    (viernes node)))
        (rows '()))
    (let* ((header-prefix (aulas-header-row-prefix-cells config))
           (column-index-labels (aulas-column-index-labels config))
           (header-row (append header-prefix column-index-labels))
           (row-length (length header-row)))
      (setf rows (list (make-blank-row row-length) header-row))
    (loop for dia in dias
          for idx from 0
          do
             (let* ((encabezados (nombres-columnas dia))
                    (nombre-dia (first encabezados))
                    (encabezados-aulas (rest encabezados))
                    (nombres (nombres-filas dia))
                    (contenido (contenido-de-la-tabla dia))
                    (row-prefix-empty (aulas-row-prefix-empty-cells config)))
               (setf rows
                     (append rows
                             (list (append (append row-prefix-empty (list nombre-dia)) encabezados-aulas))))
               (loop for i from 0 below (length nombres)
                     do
                        (setf rows
                              (append rows
                                      (list (append
                                             (append row-prefix-empty (list (nth i nombres)))
                                             (nth i contenido))))))
               (when (< idx (1- (length dias)))
                 (setf rows (append rows (list (make-blank-row row-length))))))))
    rows))

(defun python-bool (value)
  (if value "True" "False"))

(defun write-python-style (style stream)
  (format stream "{'bold': ~a, 'align': ~s, 'bg_color': '~a'}"
          (python-bool (getf style :bold))
          (getf style :align)
          (getf style :bg-color)))

(defun write-python-range-styles (styles stream)
  (format stream "    'range_styles': [~%")
  (loop for cfg in styles
        for idx from 0
        do
          (destructuring-bind (rango color) cfg
            (format stream "        {'range': '~a', 'style': {'bg_color': '~a'}}~a~%"
                    rango
                    color
                    (if (< idx (1- (length styles))) "," ""))))
  (format stream "    ],~%"))

(defmethod generate-code ((node clase-hoja)
  (lang clase-output-python-config)
  (stream t))
  (let* ((group (grupo node))
         (horario-tabla (horario node))
         (asig-tabla (asignaturas node))
         (horario-data (if horario-tabla
                           (contenido-de-la-tabla horario-tabla)
                           '()))
         (asig-data (if asig-tabla
                        (contenido-de-la-tabla asig-tabla)
                        '())))
    (format stream "grupos.append(~s)~%" group)
    (format stream "horarios_por_grupo[~s] = " group)
    (write-python-value horario-data stream)
    (format stream "~%")
    (format stream "asignaturas_por_grupo[~s] = " group)
    (write-python-value asig-data stream)
    (format stream "~2%")))

(defun construir-bloques-aulas (cantidad-turnos)
  "Construye bloques dinámicos de Aulas: (nombre fila-inicio fila-fin col-grupo fila-header)."
  (let ((dias '("Lunes" "Martes" "Miércoles" "Jueves" "Viernes"))
        (row-offset 3)
        (result '()))
    (loop for dia in dias
          for idx from 0
          for group-col-num from 3
          do
             (let* ((row-start (1+ row-offset))
                    (row-end (+ row-offset cantidad-turnos))
                    (group-col (numero-a-letra-columna group-col-num)))
               (push (list dia row-start row-end group-col row-offset) result)
               (setf row-offset (+ row-end (if (< idx (1- (length dias))) 2 0)))))
    (reverse result)))

(defun generar-formulas-aulas (grupos &key (horario-row-step 2) (cantidad-turnos 6))
  "Genera fórmulas dinámicas para la hoja Aulas.
   Cada celda cruza todas las hojas de grupo para mostrar qué grupo
   está en cada aula/turno/día."
  (let ((formulas '())
        (aula-offset (if (> horario-row-step 1) 1 0)))
    (dolist (block (construir-bloques-aulas cantidad-turnos))
      (destructuring-bind (_nombre row-start row-end group-col header-row) block
        (loop for aulas-col from 3 to 12 do
          (let* ((aulas-col-letter (numero-a-letra-columna aulas-col))
                 (header-ref (format nil "~a$~a" aulas-col-letter header-row)))
            (loop for aulas-row from row-start to row-end do
              (let* ((turno-index (- aulas-row row-start))
                     (turno-offset (if (>= turno-index 3) 1 0))
                     (group-row (+ 4 aula-offset (* turno-index horario-row-step) turno-offset))
                     (group-cell-ref (format nil "$~a$~a" group-col group-row))
                     (celda (format nil "~a~a" aulas-col-letter aulas-row))
                     (formula (generar-formula-aulas-dinamica grupos group-cell-ref header-ref)))
                (push (list :cell celda :formula formula) formulas)))))))
    (reverse formulas)))

(defun generar-formula-aulas-dinamica (grupos group-cell-ref header-ref)
  "Genera una fórmula para una celda de la hoja Aulas.
   group-cell-ref: referencia al aula en la hoja de grupo (ej: \"$C$5\")
   header-ref: referencia al header del aula en Aulas (ej: \"C$3\")"
  (let* ((fn-sustituir "SUBSTITUTE")
         (fn-espacios "TRIM")
         (fn-si "IF")
         (separador ",")
         (partes '()))
    (dolist (grupo grupos)
      (let* ((grupo-str (string grupo))
             (parte (format nil "~a(~a!~a=~a~a~a!$B$1 & \" \"~a \"\")"
                           fn-si
                           grupo-str
                           group-cell-ref
                           header-ref
                           separador
                           grupo-str
                           separador)))
        (push parte partes)))
    (format nil "=~a(~a(~a(~{~a~^,~}))~a \" \"~a \",\")"
            fn-sustituir
            fn-espacios
            "CONCAT"
            (reverse partes)
            separador
            separador
            separador)))

(defun numero-a-letra-columna (num)
  "Convierte número de columna (1-based) a letra Excel."
  (let ((resultado ""))
    (loop while (> num 0) do
      (let ((resto (mod (1- num) 26)))
        (setf resultado (concatenate 'string
                                    (string (code-char (+ 65 resto)))
                                    resultado))
        (setf num (floor (1- num) 26))))
    resultado))

(defun write-python-fernando-formulas (formulas stream)
  "Escribe las fórmulas de Aulas en formato Python."
  (format stream "    'fernando_formulas': [~%")
  (loop for f in formulas
        for idx from 0
        do
          (format stream "        {'cell': '~a', 'formula': ~s}~a~%"
                  (getf f :cell)
                  (getf f :formula)
                  (if (< idx (1- (length formulas))) "," "")))
  (format stream "    ],~%"))

(defmethod generate-code ((node clase-hoja-aulas)
                          (lang clase-output-python-config)
                          (stream t))
  (let* ((day-map (list
                    (list "Lunes" (lunes node))
                    (list "Martes" (martes node))
                    (list "Miércoles" (miercoles node))
                    (list "Jueves" (jueves node))
                    (list "Viernes" (viernes node))))
         (grupos-aulas (mapcar #'string (grupos node))))
    (dolist (entry day-map)
      (let* ((day-name (first entry))
             (tabla-dia (second entry))
             (contenido (if tabla-dia
                            (contenido-de-la-tabla tabla-dia)
                            '())))
        (format stream "aulas_por_dia[~s] = " day-name)
        (write-python-value contenido stream)
        (format stream "~%")))
    (when grupos-aulas
      (format stream "if not grupos: grupos = ")
      (write-python-value grupos-aulas stream)
      (format stream "~%"))
    (format stream "~%")))

(defmethod generate-code ((node clase-libro)
          (lang clase-output-python-config)
          (stream t))
  (format stream "from ~a import generar_excel_desde_parametros~2%"
      (python-module lang))
  (format stream "grupos = []~%")
  (format stream "horarios_por_grupo = {}~%")
  (format stream "asignaturas_por_grupo = {}~%")
  (format stream "aulas_por_dia = {}~2%")
  (loop for h in (hojas node)
    do (generate-code h lang stream))
  (format stream "generar_excel_desde_parametros(~%")
  (format stream "    filename=~s,~%" (output-excel-file lang))
  (format stream "    grupos=grupos,~%")
  (format stream "    horarios_por_grupo=horarios_por_grupo,~%")
  (format stream "    asignaturas_por_grupo=asignaturas_por_grupo,~%")
  (format stream "    aulas_por_dia=aulas_por_dia,~%")
  (format stream "    turnos=~a,~%" (or (horario-cantidad-turnos lang) 6))
  (format stream "    horario_row_step=~a~%" (max 1 (or (horario-alto-celda lang) 1)))
  (format stream ")~%"))

(defun extraer-aulas-unicas (horario-data &optional (row-step 2))
  "Extrae las aulas únicas del horario (filas pares son asignaturas, impares son aulas).
   En Lisp, los índices de lista comienzan en 0, así que:
   - La fila 0 de cada bloque es asignatura
   - La fila 1 de cada bloque es aula (si existe)"
  (let ((aulas-set (make-hash-table :test 'equal))
        (aulas-list '())
        (step (max 1 row-step))
        (aula-offset (if (> (max 1 row-step) 1) 1 0)))
    ;; Las filas de aulas son el segundo renglón de cada bloque.
    (loop for i from aula-offset below (length horario-data) by step
          do (let ((fila-aulas (nth i horario-data)))
                (loop for aula in fila-aulas
                      when (and aula (not (string= aula "")) (not (string= aula " ")))
                      do (setf (gethash aula aulas-set) t))))
    ;; Convertir a lista ordenada, primero extraer las que parecen aulas reales
    ;; Filtrar: solo incluir valores que contengan 'Aula' o 'Lab' o empiezan con minúscula
    (maphash (lambda (k v) 
               (when (or (search "Aula" k)
                         (search "Lab" k)
                         (and (> (length k) 2)
                              (char= (char k 0) #\c)
                              (char= (char k 1) #\space)))
                 (push k aulas-list)))
             aulas-set)
    (sort aulas-list #'string<)))

(defun contar-turnos-desde-datos (horario-data &optional (row-step 2))
  "Calcula la cantidad de turnos a partir de los datos del horario.
   Cada turno ocupa ROW-STEP filas.
   Se divide la longitud total entre ROW-STEP para obtener el número de turnos,
   independientemente de si las filas tienen contenido o no."
  (max 1 (floor (length horario-data) (max 1 row-step))))

(defun fila-vacia-como (ejemplo-fila)
  "Construye una fila vacía con el mismo ancho que EJEMPLO-FILA."
  (make-list (if ejemplo-fila (length ejemplo-fila) 5) :initial-element ""))

(defun normalizar-horario-data (horario-data cantidad-turnos row-step)
  "Normaliza horario-data al tamaño (cantidad-turnos * row-step).
   Si la fuente viene en formato 2 filas por turno, la expande al alto deseado."
  (let* ((step (max 1 row-step))
         (target-rows (* (max 1 cantidad-turnos) step))
         (source-rows (length horario-data))
         (sample-row (and horario-data (first horario-data))))
    (cond
      ;; Ya viene con el alto esperado.
      ((= source-rows target-rows)
       horario-data)
      ;; Caso común heredado: datos en bloques de 2 filas por turno.
      ((and (> step 2) (= source-rows (* cantidad-turnos 2)))
       (let ((result '()))
         (loop for turno from 0 below cantidad-turnos
               for base = (* turno 2)
               for asig-row = (or (nth base horario-data) (fila-vacia-como sample-row))
               for aula-row = (or (nth (1+ base) horario-data) (fila-vacia-como sample-row))
               do
                  (push asig-row result)
                  (push aula-row result)
                  (loop repeat (- step 2)
                        do (push (fila-vacia-como sample-row) result)))
         (nreverse result)))
      ;; Fallback genérico: recorta o rellena para cumplir el tamaño.
      (t
       (append (subseq horario-data 0 (min source-rows target-rows))
               (loop repeat (max 0 (- target-rows source-rows))
                     collect (fila-vacia-como sample-row)))))))

(defun crear-hoja-grupo (grupo horario-data asignaturas-data &optional (config python-config))
  (let* ((horario-row-step (max 1 (or (horario-alto-celda config) 1)))
         (cantidad-turnos (max 1 (or (horario-cantidad-turnos config)
                                     (contar-turnos-desde-datos horario-data horario-row-step))))
         (horario-data-normalizado (normalizar-horario-data horario-data cantidad-turnos horario-row-step))
         ;; Buscar la posicion de Turno 4 en los datos originales (antes de normalizar)
         ;; Esto nos da la posicion real donde empiezan los datos de Turno 4
         (turno4-start-row (loop for i from 0 below (length horario-data)
                                when (some (lambda (s) (and s (not (string= s "")))) (nth i horario-data))
                                return i))
         ;; El separador debe ir DESPUES de las filas de Turno 3
         ;; Turno 3 ocupa las filas 3*horario-row-step-1 en datos normalizados
         (separador-index (* 3 horario-row-step))
         (base-row-names (nombres-filas-por-turnos cantidad-turnos horario-row-step))
         (blank-horario-row (fila-vacia-como (and horario-data-normalizado (first horario-data-normalizado))))
         (horario-row-names
           (if (and (>= cantidad-turnos 4)
                    (<= separador-index (length base-row-names)))
               (insert-at-index base-row-names separador-index "")
               base-row-names))
         (horario-data-con-separador
           (if (and (>= cantidad-turnos 4)
                    (<= separador-index (length horario-data-normalizado)))
               (insert-at-index horario-data-normalizado separador-index blank-horario-row)
               horario-data-normalizado))
         (aulas-unicas (extraer-aulas-unicas horario-data-normalizado horario-row-step))
          ;; Crear lista de aulas lateral basada en las aulas reales del horario
          (aulas-lateral-dinamica 
             (if aulas-unicas
                 (append aulas-unicas (make-list (max 0 (- cantidad-turnos (length aulas-unicas))) :initial-element ""))
                 *aulas-lateral*)))
    (let ((horario-tabla
            (tabla-horario-aula
              :id (concatenate 'string (horario-id-prefix config) grupo)
              :cantidad-turnos cantidad-turnos
              :alto-celda (horario-alto-celda config)
              :ancho-celda (horario-ancho-celda config)
              :contenido-de-la-tabla horario-data-con-separador
              :border-color (horario-border-color config)
              :border-style (horario-border-style config)
              :range-styles (horario-range-styles config))))
      (setf (nombres-filas horario-tabla) horario-row-names)
      (setf (filas horario-tabla) (length horario-row-names))
      (hoja :grupo grupo
            :horario horario-tabla
            :asignaturas (tabla :id (concatenate 'string (asignaturas-id-prefix config) grupo)
                                :alto-celda 1
                                :ancho-celda (asignaturas-ancho-celda config)
                                :nombres-columnas (asignaturas-headers config)
                                :filas (max 1 (length asignaturas-data))
                                :contenido-de-la-tabla (if asignaturas-data asignaturas-data '()))))))

;; Los horarios ahora se cargan desde variables_horario.lisp
;; Las variables *horario-d111*, *horario-c111*, etc. están definidas allí

(defun crear-tabla-dia (dia contenido &optional (config python-config))
  (tabla :id dia
         :filas (length (aulas-row-names config))
         :columnas (1+ (length (aulas-column-names config)))
         :nombres-columnas (append (list dia) (aulas-column-names config))
         :nombres-filas (aulas-row-names config)
         :contenido-de-la-tabla contenido))

(defparameter *aulas-lunes*
  '(("" "" "" "" "" "C111,C112" "" "" "" "C113")
    ("" "" "" "" "C112" "C113" "" "" "" "")
    ("" "" "" "" "" "C111,C113" "" "C112" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")))

(defparameter *aulas-martes*
  '(("" "C112" "" "" "" "C111" "C113" "" "" "")
    ("C111" "" "" "" "" "C112,C113" "" "" "" "")
    ("" "" "" "" "" "C111,C112,C113" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")))

(defparameter *aulas-miercoles*
  '(("" "" "" "" "C112" "C111" "" "C113" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "C112" "" "" "C113" "C111" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")))

(defparameter *aulas-jueves*
  '(("C113" "" "" "" "C112" "C111" "" "" "" "")
    ("C113" "" "" "" "C112" "C111" "" "" "" "")
    ("C113" "" "" "" "C112" "C111" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")))

(defparameter *aulas-viernes*
  '(("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "C112")
    ("" "" "" "" "C111" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")
    ("" "" "" "" "" "" "" "" "" "")))

;; Crear hojas para todos los grupos desde el JSON
(defparameter *hoja-d111* (crear-hoja-grupo "D111" *horario-d111* *asignaturas-d111*))
(defparameter *hoja-d211* (crear-hoja-grupo "D211" *horario-d211* *asignaturas-d211*))
(defparameter *hoja-d311* (crear-hoja-grupo "D311" *horario-d311* *asignaturas-d311*))
(defparameter *hoja-d411* (crear-hoja-grupo "D411" *horario-d411* *asignaturas-d411*))

(defparameter *hoja-c111* (crear-hoja-grupo "C111" *horario-c111* *asignaturas-c111*))
(defparameter *hoja-c121* (crear-hoja-grupo "C121" *horario-c121* *asignaturas-c121*))
(defparameter *hoja-c122* (crear-hoja-grupo "C122" *horario-c122* *asignaturas-c122*))
(defparameter *hoja-c211* (crear-hoja-grupo "C211" *horario-c211* *asignaturas-c211*))
(defparameter *hoja-c212* (crear-hoja-grupo "C212" *horario-c212* *asignaturas-c212*))
(defparameter *hoja-c311* (crear-hoja-grupo "C311" *horario-c311* *asignaturas-c311*))
(defparameter *hoja-c312* (crear-hoja-grupo "C312" *horario-c312* *asignaturas-c312*))
(defparameter *hoja-c411* (crear-hoja-grupo "C411" *horario-c411* *asignaturas-c411*))
(defparameter *hoja-c412* (crear-hoja-grupo "C412" *horario-c412* *asignaturas-c412*))

(defparameter *hoja-m111* (crear-hoja-grupo "M111" *horario-m111* *asignaturas-m111*))
(defparameter *hoja-m211* (crear-hoja-grupo "M211" *horario-m211* *asignaturas-m211*))
(defparameter *hoja-m311* (crear-hoja-grupo "M311" *horario-m311* *asignaturas-m311*))
(defparameter *hoja-m411* (crear-hoja-grupo "M411" *horario-m411* *asignaturas-m411*))

(defparameter *grupos-todos*
  '(D111 D211 D311 D411
    C111 C121 C122
    C211 C212
    C311 C312
    C411 C412
    M111 M211 M311 M411))

(defparameter *hoja-aulas*
  (hoja-aulas :lunes (crear-tabla-dia "Lunes" *aulas-lunes*)
              :martes (crear-tabla-dia "Martes" *aulas-martes*)
              :miercoles (crear-tabla-dia "Miércoles" *aulas-miercoles*)
              :jueves (crear-tabla-dia "Jueves" *aulas-jueves*)
              :viernes (crear-tabla-dia "Viernes" *aulas-viernes*)
              :grupos *grupos-todos*))

(defparameter *libro-propuesta*
  (libro :hojas (list *hoja-d111* *hoja-d211* *hoja-d311* *hoja-d411*
                       *hoja-c111* *hoja-c121* *hoja-c122* 
                       *hoja-c211* *hoja-c212* 
                       *hoja-c311* *hoja-c312* 
                       *hoja-c411* *hoja-c412*
                       *hoja-m111* *hoja-m211* *hoja-m311* *hoja-m411*
                       *hoja-aulas*)))

(with-open-file (f "generar_propuesta_desde_lisp.py"
                   :direction :output
                   :if-exists :supersede)
  (generate-code *libro-propuesta* python-config f))

(format t "Archivo Python 'generar_propuesta_desde_lisp.py' generado con exito.~%")
